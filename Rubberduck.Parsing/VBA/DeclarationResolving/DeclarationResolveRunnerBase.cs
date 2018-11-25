using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.ComReferenceLoading;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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

                var annotations = _state.GetModuleAnnotations(module).ToList();
                var attributes = _state.GetModuleAttributes(module);
                var membersAllowingAttributes = _state.GetMembersAllowingAttributes(module);

                var moduleDeclaration = NewModuleDeclaration(module, tree, annotations, attributes, projectDeclaration);
                _state.AddDeclaration(moduleDeclaration);

                var controlDeclarations = DeclarationsFromControls(moduleDeclaration);
                foreach (var declaration in controlDeclarations)
                {
                    _state.AddDeclaration(declaration);
                }

                var declarationsListener = new DeclarationSymbolsListener(moduleDeclaration, annotations, attributes, membersAllowingAttributes);
                ParseTreeWalker.Default.Walk(declarationsListener, tree);
                foreach (var createdDeclaration in declarationsListener.CreatedDeclarations)
                {
                    _state.AddDeclaration(createdDeclaration);
                }

                //This is a hack to deal with annotations on module level variables.
                var memberAnnotations = declarationsListener.CreatedDeclarations
                    .SelectMany(declaration => declaration.Annotations)
                    .ToHashSet();
                moduleDeclaration.RemoveAnnotations(memberAnnotations);
            }
            catch (Exception exception)
            {
                Logger.Error(exception, $"Exception thrown acquiring declarations for '{module.Name}' (thread {Thread.CurrentThread.ManagedThreadId}).");
                _parserStateManager.SetModuleState(module, ParserState.ResolverError, token);
            }
            stopwatch.Stop();
            Logger.Debug($"{stopwatch.ElapsedMilliseconds}ms to resolve declarations for component {module.Name}");
        }

        private ModuleDeclaration NewModuleDeclaration(
            QualifiedModuleName qualifiedModuleName,
            IParseTree tree,
            ICollection<IAnnotation> annotations,
            IDictionary<(string scopeIdentifier, DeclarationType scopeType),
                Attributes> attributes,
            Declaration projectDeclaration)
        {
            var moduleAttributes = ModuleAttributes(qualifiedModuleName, attributes);
            var moduleAnnotations = FindModuleAnnotations(tree, annotations);

            switch (qualifiedModuleName.ComponentType)
            {
                case ComponentType.StandardModule:
                    return new ProceduralModuleDeclaration(
                        qualifiedModuleName.QualifyMemberName(qualifiedModuleName.ComponentName),
                        projectDeclaration,
                        qualifiedModuleName.ComponentName,
                        true,
                        moduleAnnotations,
                        moduleAttributes);
                case ComponentType.ClassModule:
                    return new ClassModuleDeclaration(
                        qualifiedModuleName.QualifyMemberName(qualifiedModuleName.ComponentName),
                        projectDeclaration,
                        qualifiedModuleName.ComponentName,
                        true,
                        moduleAnnotations,
                        moduleAttributes);
                default:
                    return new ClassModuleDeclaration(
                        qualifiedModuleName.QualifyMemberName(qualifiedModuleName.ComponentName),
                        projectDeclaration,
                        qualifiedModuleName.ComponentName,
                        true,
                        moduleAnnotations,
                        moduleAttributes,
                        hasDefaultInstanceVariable: true);
            }
        }

        private static Attributes ModuleAttributes(QualifiedModuleName qualifiedModuleName, IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> attributes)
        {
            var declarationType = qualifiedModuleName.ComponentType == ComponentType.StandardModule
                ? DeclarationType.ProceduralModule
                : DeclarationType.ClassModule;
            var key = (qualifiedModuleName.ComponentName, declarationType);
            var moduleAttributes = attributes.ContainsKey(key)
                ? attributes[key]
                : new Attributes();
            return moduleAttributes;
        }

        private static IEnumerable<IAnnotation> FindModuleAnnotations(IParseTree tree, ICollection<IAnnotation> annotations)
        {
            if (annotations == null)
            {
                return null;
            }

            var lastDeclarationsSectionLine = LastDeclarationsSectionLine(tree, annotations);

            //There is no module body.
            if (lastDeclarationsSectionLine == null)
            {
                return annotations;
            }

            var lastPossibleModuleAnnotationLine = lastDeclarationsSectionLine.Value;
            var moduleAnnotations = annotations.Where(annotation => annotation.QualifiedSelection.Selection.EndLine <= lastPossibleModuleAnnotationLine);
            return moduleAnnotations.ToList();
        }

        private static int? LastDeclarationsSectionLine(IParseTree tree, ICollection<IAnnotation> annotations)
        {
            var firstModuleBodyElementLine = FirstModuleBodyElementLine(tree);

            if (firstModuleBodyElementLine == null)
            {
                return null;
            }

            //The VBE uses 1-based lines.
            for (var currentLine = firstModuleBodyElementLine.Value - 1; currentLine >= 1; currentLine--)
            {
                if (annotations.Any(annotation => annotation.QualifiedSelection.Selection.StartLine <= currentLine
                                                   && annotation.QualifiedSelection.Selection.EndLine >=
                                                   currentLine))
                {
                    continue;
                }

                return currentLine;
            }

            //There is no declaration section.
            return 0;
        }

        private static int? FirstModuleBodyElementLine(IParseTree tree)
        {
            var startContext = (ParserRuleContext)tree;
            var moduleBody = startContext.GetDescendent<VBAParser.ModuleBodyContext>();

            var moduleBodyElements = moduleBody.moduleBodyElement();
            if (!moduleBodyElements.Any())
            {
                return null;
            }

            return moduleBodyElements.Select(context => context.start.Line).Min();
        }

        /// <summary>
        /// Scans form designer to create a public, self-assigned field for each control on a form.
        /// </summary>
        /// <remarks>
        /// These declarations are meant to be used to identify control event procedures.
        /// </remarks>
        private IEnumerable<Declaration> DeclarationsFromControls(Declaration moduleDeclaration)
        {
            var module = moduleDeclaration.QualifiedModuleName;
            var form = _state.ProjectsProvider.Component(module);

            if (form == null || !(module.ComponentType == ComponentType.UserForm || form.HasDesigner))
            {
                return new List<Declaration>();
            }

            using (var controls = form.Controls)
            {
                if (controls == null)
                {
                    return new List<Declaration>();
                }

                var libraryQualifier = string.Empty;
                if (module.ComponentType == ComponentType.UserForm)
                {
                    var msFormsLib = _state.DeclarationFinder.FindProject("MSForms");
                    if (msFormsLib != null)
                    {
                        // given a UserForm component, MSForms reference is in use and cannot be removed.
                        libraryQualifier = "MSForms.";
                    }
                }

                var controlDeclarations = new List<Declaration>();
                foreach (var control in controls)
                {
                    var typeName = $"{libraryQualifier}{control.TypeName()}";
                    // The as type declaration should be TextBox, CheckBox, etc. depending on the type.
                    var declaration = new Declaration(
                        module.QualifyMemberName(control.Name),
                        moduleDeclaration,
                        moduleDeclaration,
                        string.IsNullOrEmpty(typeName) ? "Control" : typeName,
                        null,
                        true,
                        true,
                        Accessibility.Public,
                        DeclarationType.Control,
                        null,
                        null,
                        Selection.Home,
                        false,
                        null,
                        true);

                    controlDeclarations.Add(declaration);
                }

                return controlDeclarations;
            }
        }
    }
}
