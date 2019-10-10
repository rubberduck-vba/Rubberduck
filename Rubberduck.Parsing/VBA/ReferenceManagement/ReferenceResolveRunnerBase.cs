﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.ReferenceManagement.CompilationPasses;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
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

            var parsingStageTimer = ParsingStageTimer.StartNew();

            ExecuteCompilationPasses(_toResolve.AsReadOnly(), token);
            token.ThrowIfCancellationRequested();

            parsingStageTimer.Stop();
            parsingStageTimer.Log("Executed compilation passes in {0}ms.");
            parsingStageTimer.Restart();

            AddSupertypesForDocumentModules(_toResolve.AsReadOnly(), _state);
            token.ThrowIfCancellationRequested();

            parsingStageTimer.Stop();
            parsingStageTimer.Log("Added supertypes for document modules in {0}ms.");

            var parseTreesToResolve = _state.ParseTrees.Where(kvp => _toResolve.Contains(kvp.Key)).ToList();
            token.ThrowIfCancellationRequested();

            parsingStageTimer.Restart();

            ResolveReferences(parseTreesToResolve, token);
            token.ThrowIfCancellationRequested();

            parsingStageTimer.Stop();
            parsingStageTimer.Log("Resolved references in {0}ms.");
            parsingStageTimer.Restart();

            AddModuleToModuleReferences(_state.DeclarationFinder, token);
            token.ThrowIfCancellationRequested();

            parsingStageTimer.Stop();
            parsingStageTimer.Log("Determined module to module references in {0}ms.");
            parsingStageTimer.Restart();

            AddNewUndeclaredVariablesToDeclarations();
            AddNewFailedResolutions();

            _toResolve.Clear();
        }

        private void PerformPreResolveCleanup(IReadOnlyCollection<QualifiedModuleName> toResolve, CancellationToken token)
        {
            _referenceRemover.RemoveReferencesBy(toResolve, token);
            _moduleToModuleReferenceManager.ClearModuleToModuleReferencesFromModule(toResolve);
            _moduleToModuleReferenceManager.ClearModuleToModuleReferencesToModule(toResolve);
            foreach (var module in toResolve)
            {
                _state.ClearFailedResolutions(module);
            }
        }

        private void ExecuteCompilationPasses(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            var passes = new List<ICompilationPass>
                {
                    // This pass has to come first because the type binding resolution depends on it.
                    new ProjectReferencePass(_state.DeclarationFinder),
                    new TypeHierarchyPass(_state.DeclarationFinder, new VBAExpressionParser()),
                    new TypeAnnotationPass(_state.DeclarationFinder, new VBAExpressionParser())
                };
            try
            {
                passes.ForEach(p => p.Execute(modules));
            }
            catch (Exception exception)
            {
                var names = string.Join(",", modules.Select(m => m.Name));
                Logger.Error(exception, "Exception thrown on resolving those modules: '{0}' (thread {1}).", names, Thread.CurrentThread.ManagedThreadId);
                _parserStateManager.SetModuleStates(modules, ParserState.ResolverError, token);
            }
        }

        private void AddSupertypesForDocumentModules(IReadOnlyCollection<QualifiedModuleName> modules, RubberduckParserState state)
        {
            var documentModuleDeclarations = state.DeclarationFinder.UserDeclarations(DeclarationType.Document)
                .OfType<DocumentModuleDeclaration>()
                .Where(declaration => modules.Contains(declaration.QualifiedName.QualifiedModuleName));

            foreach (var documentDeclaration in documentModuleDeclarations)
            {
                var documentSupertype = SupertypeForDocument(documentDeclaration.QualifiedName.QualifiedModuleName, state);
                if (documentSupertype != null)
                {
                    documentDeclaration.AddSupertype(documentSupertype);
                }
            }
        }

        private Declaration SupertypeForDocument(QualifiedModuleName module, RubberduckParserState state)
        {
            if(module.ComponentType != ComponentType.Document)
            {
                return null;
            }

            var component = _state.ProjectsProvider.Component(module);
            if (component == null || component.IsWrappingNullReference)
            {
                return null;
            }

            Declaration superType = null;
            // TODO: Replace with TypeLibAPI call, require a solution regarding thread synchronization or caching
            /*
            using (var properties = component.Properties)
            {
                int documentPropertyCount = 0;
                try
                {
                    if (properties == null || properties.IsWrappingNullReference)
                    {
                        return null;
                    }
                    documentPropertyCount = properties.Count;
                }
                catch(COMException)
                {
                    return null;
                }
                
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
                            using (var property = properties[i+1])
                            {
                                if (coclass.Key[i] != property?.Name)
                                {
                                    allNamesMatch = false;
                                    break;
                                }
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
            }
            */

            return superType;
        }

        protected void ResolveReferences(DeclarationFinder finder, QualifiedModuleName module, IParseTree tree, CancellationToken token)
        {
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

        private void AddNewFailedResolutions()
        {
            var failedResolutionStores = _state.DeclarationFinder.FreshFailedResolutionStores;
            foreach (var (module, store) in failedResolutionStores)
            {
                _state.AddFailedResolutions(module, store);
            }
        }
    }
}
