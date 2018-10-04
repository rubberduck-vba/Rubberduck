using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.ComReferenceLoading;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Parsing.VBA.DeclarationResolving
{
    public abstract class DeclarationResolveRunnerBase : IDeclarationResolveRunner
    {
        protected static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        protected readonly RubberduckParserState _state;
        protected readonly IParserStateManager _parserStateManager;
        private readonly IProjectReferencesProvider _projectReferencesProvider;

        protected DeclarationResolveRunnerBase(
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


        public void CreateProjectDeclarations(IReadOnlyCollection<string> projectIds)
        {
            var existingProjectDeclarations = ProjectDeclarations();
            foreach (var projectId in projectIds)
            {
                if (existingProjectDeclarations.ContainsKey(projectId))
                {
                    continue;
                }

                var projectDeclaration = CreateProjectDeclaration(projectId);
                _state.AddDeclaration(projectDeclaration);
            }
        }

        private IDictionary<string, ProjectDeclaration> ProjectDeclarations()
        {
            var projectDeclarations = _state.DeclarationFinder
                .UserDeclarations(DeclarationType.Project)
                .Cast<ProjectDeclaration>()
                .ToDictionary(project => project.ProjectId);
            return projectDeclarations;
        }

        private Declaration CreateProjectDeclaration(string projectId)
        {
            var project = _state.ProjectsProvider.Project(projectId);

            var qualifiedModuleName = new QualifiedModuleName(project);
            var qualifiedName = qualifiedModuleName.QualifyMemberName(project.Name);
            var projectDeclaration = new ProjectDeclaration(qualifiedName, qualifiedName.MemberName, true, project);

            return projectDeclaration;
        }

        public void RefreshProjectReferences()
        {
            var existingProjects = ProjectDeclarations();
            foreach (var (projectId, projectDeclaration) in existingProjects)
            {
                projectDeclaration.ClearProjectReferences();
                var references = ProjectReferences(projectId);
                AddReferences(projectDeclaration, references);
            }
        }

        private static void AddReferences(ProjectDeclaration projectDeclaration, List<ReferencePriorityMap> references)
        {
            var projectId = projectDeclaration.ProjectId;
            foreach (var reference in references)
            {
                int priority = reference[projectId];
                projectDeclaration.AddProjectReference(reference.ReferencedProjectId, priority);
            }
        }

        private List<ReferencePriorityMap> ProjectReferences(string projectId)
        {
            var references = new List<ReferencePriorityMap>();
            foreach (var item in _projectReferencesProvider.ProjectReferences)
            {
                if (item.ContainsKey(projectId))
                {
                    references.Add(item);
                }
            }

            return references;
        }

        public void ResolveDeclarations(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            var projectDeclarations = ProjectDeclarations();
            ResolveDeclarations(modules, projectDeclarations, token);
        }

        protected abstract void ResolveDeclarations(IReadOnlyCollection<QualifiedModuleName> modules, IDictionary<string, ProjectDeclaration> projects, CancellationToken token);

        protected void ResolveDeclarations(QualifiedModuleName module, IParseTree tree, IDictionary<string, ProjectDeclaration> projects, CancellationToken token)
        {
            var stopwatch = Stopwatch.StartNew();
            try
            {
                if (!projects.TryGetValue(module.ProjectId, out var projectDeclaration))
                {
                    Logger.Error($"Tried to add module {module} with projectId {module.ProjectId} for which no project declaration exists.");
                }
                Logger.Debug($"Creating declarations for module {module.Name}.");

                var declarationsListener = new DeclarationSymbolsListener(_state, module, _state.GetModuleAnnotations(module), _state.GetModuleAttributes(module), _state.GetMembersAllowingAttributes(module), projectDeclaration);
                ParseTreeWalker.Default.Walk(declarationsListener, tree);
                foreach (var createdDeclaration in declarationsListener.CreatedDeclarations)
                {
                    _state.AddDeclaration(createdDeclaration);
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, $"Exception thrown acquiring declarations for '{module.Name}' (thread {Thread.CurrentThread.ManagedThreadId}).");
                _parserStateManager.SetModuleState(module, ParserState.ResolverError, token);
            }
            stopwatch.Stop();
            Logger.Debug($"{stopwatch.ElapsedMilliseconds}ms to resolve declarations for component {module.Name}");
        }
    }
}
