using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Symbols;
using Antlr4.Runtime.Tree;
using System.Diagnostics;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Runtime.InteropServices;

namespace Rubberduck.Parsing.VBA
{
    public abstract class ReferenceResolveRunnerBase : IReferenceResolveRunner
    {
        //This is used to buffer the modules for which to resolve references
        //to avoid missing modules in case of a cancellation or a resolver error.
        private readonly HashSet<QualifiedModuleName> _toResolve = new HashSet<QualifiedModuleName>();

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        protected readonly RubberduckParserState _state;
        protected readonly IParserStateManager _parserStateManager;
        private readonly IModuleToModuleReferenceManager _moduleToModuleReferenceManager;
        private readonly IReferenceRemover _referenceRemover;

        public ReferenceResolveRunnerBase(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            IModuleToModuleReferenceManager moduletToModuleReferenceManager,
            IReferenceRemover referenceRemover)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }
            if (parserStateManager == null)
            {
                throw new ArgumentNullException(nameof(parserStateManager));
            }
            if (moduletToModuleReferenceManager == null)
            {
                throw new ArgumentNullException(nameof(moduletToModuleReferenceManager));
            }
            if (referenceRemover == null)
            {
                throw new ArgumentNullException(nameof(referenceRemover));
            }

            _state = state;
            _parserStateManager = parserStateManager;
            _moduleToModuleReferenceManager = moduletToModuleReferenceManager;
            _referenceRemover = referenceRemover;
        }


        protected abstract void ResolveReferences(ICollection<KeyValuePair<QualifiedModuleName, IParseTree>> toResolve, CancellationToken token);
        protected abstract void AddModuleToModuleReferences(DeclarationFinder finder, CancellationToken token);


        public void ResolveReferences(IReadOnlyCollection<QualifiedModuleName> toResolve, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            _toResolve.UnionWith(toResolve);
            token.ThrowIfCancellationRequested();

            if(!_toResolve.Any())
            {
                return;
            }

            PerformPreResolveCleanup(_toResolve.AsReadOnly(), token);
            token.ThrowIfCancellationRequested();

            ExecuteCompilationPasses(_toResolve.AsReadOnly());
            token.ThrowIfCancellationRequested();

            AddSupertypesForDocumentModules(_toResolve.AsReadOnly(), _state);
            token.ThrowIfCancellationRequested();

            var parseTreesToResolve = _state.ParseTrees.Where(kvp => _toResolve.Contains(kvp.Key)).ToList();
            token.ThrowIfCancellationRequested();

            ResolveReferences(parseTreesToResolve, token);
            token.ThrowIfCancellationRequested();

            AddModuleToModuleReferences(_state.DeclarationFinder, token);
            token.ThrowIfCancellationRequested();

            AddNewUndeclaredVariablesToDeclarations();
            AddNewUnresolvedMemberDeclarations();

            _toResolve.Clear();
        }

        private void PerformPreResolveCleanup(IReadOnlyCollection<QualifiedModuleName> toResolve, CancellationToken token)
        {
            _referenceRemover.RemoveReferencesBy(toResolve, token);
            _moduleToModuleReferenceManager.ClearModuleToModuleReferencesFromModule(toResolve);
            _moduleToModuleReferenceManager.ClearModuleToModuleReferencesToModule(toResolve);
        }

        private void ExecuteCompilationPasses(IReadOnlyCollection<QualifiedModuleName> modules)
        {
            var passes = new List<ICompilationPass>
                {
                    // This pass has to come first because the type binding resolution depends on it.
                    new ProjectReferencePass(_state.DeclarationFinder),
                    new TypeHierarchyPass(_state.DeclarationFinder, new VBAExpressionParser()),
                    new TypeAnnotationPass(_state.DeclarationFinder, new VBAExpressionParser())
                };
            passes.ForEach(p => p.Execute(modules));
        }

        private void AddSupertypesForDocumentModules(IReadOnlyCollection<QualifiedModuleName> modules, RubberduckParserState state)
        {
            var allClassModuleDeclarations = state.DeclarationFinder.UserDeclarations(DeclarationType.ClassModule);
            var documentModuleDeclarations = allClassModuleDeclarations.Where(declaration =>
                                                                                declaration.QualifiedName.QualifiedModuleName.ComponentType == ComponentType.Document
                                                                                && modules.Contains(declaration.QualifiedName.QualifiedModuleName));

            foreach (var documentDeclaration in documentModuleDeclarations)
            {
                var documentSupertype = SupertypeForDocument(documentDeclaration.QualifiedName.QualifiedModuleName, state);
                if (documentSupertype != null)
                {
                    ((ClassModuleDeclaration)documentDeclaration).AddSupertype(documentSupertype);
                }
            }
        }

        private Declaration SupertypeForDocument(QualifiedModuleName module, RubberduckParserState state)
        {
            if(module.ComponentType != ComponentType.Document || module.Component == null)
            {
                return null;
            }

            int documentPropertyCount = 0;
            try
            {
                if(module.Component.IsWrappingNullReference
                    || module.Component.Properties == null
                    || module.Component.Properties.IsWrappingNullReference)
                {
                    return null;
                }
                documentPropertyCount = module.Component.Properties.Count;
            }
            catch(COMException)
            {
                return null;
            }

            Declaration superType = null;
            foreach (var coclass in state.CoClasses)
            {
                try
                {
                    

                    if (coclass.Key.Count != documentPropertyCount)
                    {
                        continue;
                    }

                    var allNamesMatch = true;
                    for (var i = 0; i < coclass.Key.Count; i++)
                    {
                        if (coclass.Key[i] != module.Component.Properties[i + 1].Name)
                        {
                            allNamesMatch = false;
                            break;
                        }
                    }

                    if (allNamesMatch)
                    {
                        superType = coclass.Value;
                        break;
                    }
                }
                catch (COMException)
                {
                }
            }

            return superType;
        }

        protected void ResolveReferences(DeclarationFinder finder, QualifiedModuleName module, IParseTree tree, CancellationToken token)
        {
            Debug.Assert(_state.GetModuleState(module) == ParserState.ResolvingReferences || token.IsCancellationRequested);

            token.ThrowIfCancellationRequested();

            Logger.Debug("Resolving identifier references in '{0}'... (thread {1})", module.Name, Thread.CurrentThread.ManagedThreadId);

            var resolver = new IdentifierReferenceResolver(module, finder);
            var listener = new IdentifierReferenceListener(resolver);

            if (!string.IsNullOrWhiteSpace(tree.GetText().Trim()))
            {
                var walker = new ParseTreeWalker();
                try
                {
                    var watch = Stopwatch.StartNew();
                    walker.Walk(listener, tree);
                    watch.Stop();
                    Logger.Debug("Binding resolution done for component '{0}' in {1}ms (thread {2})", module.Name,
                        watch.ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);

                    //Evaluation of the overall status has to be defered to allow processing of undeclared variables before setting the ready state.
                    _parserStateManager.SetModuleState(module, ParserState.Ready, token, false);
                }
                catch (OperationCanceledException)
                {
                    throw;  //We do not want to set an error state if the exception was just caused by some cancellation.
                }
                catch (Exception exception)
                {
                    Logger.Error(exception, "Exception thrown resolving '{0}' (thread {1}).", module.Name, Thread.CurrentThread.ManagedThreadId);
                    _parserStateManager.SetModuleState(module, ParserState.ResolverError, token);
                }
            }
        }

        protected void AddModuleToModuleReferences(DeclarationFinder finder, QualifiedModuleName referencedModule)
        {
            var referencingModules = finder.Members(referencedModule)
                                        .SelectMany(declaration => declaration.References)
                                        .Select(reference => reference.QualifiedModuleName)
                                        .Distinct()
                                        .Where(referencingModule => !referencedModule.Equals(referencingModule));
            foreach (var referencingModule in referencingModules)
            {
                _moduleToModuleReferenceManager.AddModuleToModuleReference(referencingModule, referencedModule);
            }
        }

        private void AddNewUndeclaredVariablesToDeclarations()
        {
            var undeclared = _state.DeclarationFinder.FreshUndeclared;
            foreach (var declaration in undeclared)
            {
                _state.AddDeclaration(declaration);
            }
        }

        private void AddNewUnresolvedMemberDeclarations()
        {
            var unresolved = _state.DeclarationFinder.FreshUnresolvedMemberDeclarations;
            foreach (var declaration in unresolved)
            {
                _state.AddUnresolvedMemberDeclaration(declaration);
            }
        }
    }
}
