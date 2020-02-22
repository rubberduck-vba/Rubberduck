using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerProjectViewModel : CodeExplorerItemViewModel
    {
        public static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.ClassModule,
            DeclarationType.Document,
            DeclarationType.ProceduralModule,
            DeclarationType.UserForm
        };

        private readonly IVBE _vbe;
        private readonly IProjectsProvider _projectsProvider;

        public CodeExplorerProjectViewModel(
            Declaration project, 
            ref List<Declaration> declarations, 
            RubberduckParserState state, 
            IVBE vbe, 
            IProjectsProvider projectsProvider,
            bool references = true) 
            : base(null, project)
        {
            State = state;
            _vbe = vbe;
            _projectsProvider = projectsProvider;
            ShowReferences = references;

            SetName();
            var children = ExtractTrackedDeclarationsForProject(project, ref declarations);
            AddNewChildren(ref children);
            IsExpanded = true;
        }

        private string _displayName;
        private string _name;

        public RubberduckParserState State { get; }

        public bool ShowReferences { get; }

        public override string Name => string.IsNullOrEmpty(_displayName) ? _name : $"{_name} ({_displayName})";
        
        public override FontWeight FontWeight
        {
            get
            {
                if (_vbe.Kind == VBEKind.Hosted || Declaration == null)
                {
                    return base.FontWeight;
                }

                var project = _projectsProvider.Project(Declaration.ProjectId);
                if (project == null)
                {
                    return base.FontWeight;
                }

                using (var vbProjects = _vbe.VBProjects)
                using (var startProject = vbProjects?.StartProject)
                {
                    return project.Equals(startProject) 
                        ? FontWeights.Bold 
                        : base.FontWeight;
                }
            }
        }

        public override Comparer<ICodeExplorerNode> SortComparer => CodeExplorerItemComparer.NodeType;

        public override bool Filtered => false;

        public override void Synchronize(ref List<Declaration> updated)
        {
            if (Declaration is null ||
                !(updated?.OfType<ProjectDeclaration>()
                    .FirstOrDefault(declaration => declaration.ProjectId.Equals(Declaration.ProjectId)) is ProjectDeclaration match))
            {
                Declaration = null;
                return;
            }

            Declaration = match;

            var children = ExtractTrackedDeclarationsForProject(Declaration, ref updated);
            updated = updated.Except(children.Union(new[] { Declaration })).ToList();

            // Reference synchronization is deferred to AddNewChildren for 2 reasons. First, it doesn't make sense to sling around a List of
            // declaration for something that doesn't need it. Second (and more importantly), the priority can't be set without calling 
            // GetProjectReferenceModels, which hits the VBE COM interfaces. So, we only want to do that once. The bonus 3rd reason is that it
            // can be called from the ctor this way.

            SynchronizeChildren(ref children);

            // Have to do this again - the project might have been saved or otherwise had the ProjectDisplayName changed.
            SetName();
        }

        protected sealed override void AddNewChildren(ref List<Declaration> updated)
        {
            if (updated is null)
            {
                return;
            }

            SynchronizeReferences();

            foreach (var rootFolder in updated.GroupBy(declaration => declaration.RootFolder())
                .Where(folder => !string.IsNullOrEmpty(folder.Key)))
            {
                var contents = rootFolder.ToList();
                AddChild(new CodeExplorerCustomFolderViewModel(this, rootFolder.Key, rootFolder.Key, _vbe, ref contents));
            }
        }

        private void SynchronizeReferences()
        {
            if (!ShowReferences)
            {
                return;
            }

            var references = GetProjectReferenceModels();
            foreach (var child in Children.OfType<CodeExplorerReferenceFolderViewModel>())
            {
                child.Synchronize(Declaration, references);
                if (child.Declaration is null)
                {
                    RemoveChild(child);
                }
            }

            if (!references.Any())
            {
                return;
            }

            var types = references.GroupBy(reference => reference.Type);

            foreach (var type in types)
            {
                AddChild(new CodeExplorerReferenceFolderViewModel(this, State?.DeclarationFinder, type.ToList(), type.Key));
            }
        }

        private List<ReferenceModel> GetProjectReferenceModels()
        {
            if (Declaration == null)
            {
                return new List<ReferenceModel>();
            }

            var project = _projectsProvider.Project(Declaration.ProjectId);
            if (project == null)
            {
                return new List<ReferenceModel>();
            }

            var referenced = new List<ReferenceModel>();

            using (var references = project.References)
            {
                var priority = 1;
                foreach (var reference in references)
                {
                    referenced.Add(new ReferenceModel(reference, priority++));
                    reference.Dispose();
                }
            }

            return referenced;
        }

        private void SetName()
        {
            if (Declaration is null)
            {
                return;
            }

            _name = Declaration?.IdentifierName ?? string.Empty;

            // F' the flicker. Digging into the properties has some even more evil side-effects, and is a performance nightmare by comparison.
            _displayName = DisplayName(Declaration);

            OnNameChanged();
        }

        private string DisplayName(Declaration declaration)
        {
            if (declaration == null)
            {
                return string.Empty;
            }

            var project = _projectsProvider.Project(declaration.ProjectId);
            return project != null 
                ? project.ProjectDisplayName 
                : string.Empty;
        }

        private static readonly List<DeclarationType> UntrackedTypes = new List<DeclarationType>
        {
            DeclarationType.Parameter,
            DeclarationType.LineLabel,
            DeclarationType.UnresolvedMember,
            DeclarationType.BracketedExpression,
            DeclarationType.ComAlias
        };

        private static readonly List<DeclarationType> ModuleRestrictedTypes = new List<DeclarationType>
        {
            DeclarationType.Variable,
            DeclarationType.Control,
            DeclarationType.Constant
        };

        public static List<Declaration> ExtractTrackedDeclarationsForProject(Declaration project, ref List<Declaration> declarations)
        {
            var owned = declarations.Where(declaration => declaration.ProjectId.Equals(project.ProjectId)).ToList();
            declarations = declarations.Except(owned).ToList();

            return owned.Where(declaration => !UntrackedTypes.Contains(declaration.DeclarationType) &&
                               (!ModuleRestrictedTypes.Contains(declaration.DeclarationType) ||
                                declaration.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module))).ToList();
        }
    }
}
