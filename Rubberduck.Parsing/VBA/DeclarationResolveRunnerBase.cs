using System;
using System.Collections.Generic;
using System.Threading;
using Rubberduck.VBEditor;
using System.Collections.Concurrent;
using Rubberduck.Parsing.Symbols;
using Antlr4.Runtime.Tree;
using System.Diagnostics;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public abstract class DeclarationResolveRunnerBase : IDeclarationResolveRunner
    {
        protected readonly ConcurrentDictionary<string, Declaration> _projectDeclarations = new ConcurrentDictionary<string, Declaration>();
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        protected readonly RubberduckParserState _state;
        protected readonly IParserStateManager _parserStateManager;
        private readonly IProjectReferencesProvider _projectReferencesProvider;

        public DeclarationResolveRunnerBase(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            IProjectReferencesProvider projectReferencesProvider)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }
            if (parserStateManager == null)
            {
                throw new ArgumentNullException(nameof(parserStateManager));
            }
            if (projectReferencesProvider == null)
            {
                throw new ArgumentNullException(nameof(projectReferencesProvider));
            }

            _state = state;
            _parserStateManager = parserStateManager;
            _projectReferencesProvider = projectReferencesProvider;
        }


        public abstract void ResolveDeclarations(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token);


        protected void ResolveDeclarations(QualifiedModuleName module, IParseTree tree, CancellationToken token)
        {
            if (module == null) { return; }

            var stopwatch = Stopwatch.StartNew();
            try
            {
                var project = module.Component.Collection.Parent;
                var projectQualifiedName = new QualifiedModuleName(project);
                Declaration projectDeclaration;
                if (!_projectDeclarations.TryGetValue(projectQualifiedName.ProjectId, out projectDeclaration))
                {
                    projectDeclaration = CreateProjectDeclaration(projectQualifiedName, project);
                    _projectDeclarations.AddOrUpdate(projectQualifiedName.ProjectId, projectDeclaration, (s, c) => projectDeclaration);
                    _state.AddDeclaration(projectDeclaration);
                }
                Logger.Debug("Creating declarations for module {0}.", module.Name);

                var declarationsListener = new DeclarationSymbolsListener(_state, module, module.ComponentType, _state.GetModuleAnnotations(module), _state.GetModuleAttributes(module), projectDeclaration);
                ParseTreeWalker.Default.Walk(declarationsListener, tree);
                foreach (var createdDeclaration in declarationsListener.CreatedDeclarations)
                {
                    _state.AddDeclaration(createdDeclaration);
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Exception thrown acquiring declarations for '{0}' (thread {1}).", module.Name, Thread.CurrentThread.ManagedThreadId);
                _parserStateManager.SetModuleState(module, ParserState.ResolverError, token);
            }
            stopwatch.Stop();
            Logger.Debug("{0}ms to resolve declarations for component {1}", stopwatch.ElapsedMilliseconds, module.Name);
        }

        private Declaration CreateProjectDeclaration(QualifiedModuleName projectQualifiedName, IVBProject project)
        {
            var qualifiedName = projectQualifiedName.QualifyMemberName(project.Name);
            var projectId = qualifiedName.QualifiedModuleName.ProjectId;
            var projectDeclaration = new ProjectDeclaration(qualifiedName, project.Name, true, project);

            var references = new List<ReferencePriorityMap>();
            foreach (var item in _projectReferencesProvider.ProjectReferences)
            {
                if (item.ContainsKey(projectId))
                {
                    references.Add(item);
                }
            }

            foreach (var reference in references)
            {
                int priority = reference[projectId];
                projectDeclaration.AddProjectReference(reference.ReferencedProjectId, priority);
            }
            return projectDeclaration;
        }
    }
}
