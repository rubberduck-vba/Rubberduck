using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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

        public CodeExplorerProjectViewModel(Declaration declaration, IEnumerable<Declaration> declarations, RubberduckParserState state, IVBE vbe) : base(null, declaration)
        {
            State = state;
            _vbe = vbe;
            SetName();
            AddNewChildren(declarations.ToList());
        }

        private string _displayName;
        private string _name;

        public override string Name => string.IsNullOrEmpty(_displayName) ? _name : $"{_name} ({_displayName})";

        public RubberduckParserState State { get; }

        public override FontWeight FontWeight
        {
            get
            {
                if (_vbe.Kind == VBEKind.Hosted || Declaration.Project == null)
                {
                    return base.FontWeight;
                }

                using (var vbProjects = _vbe.VBProjects)
                using (var startProject = vbProjects?.StartProject)
                {
                    return Declaration.Project.Equals(startProject) ? FontWeights.Bold : base.FontWeight;
                }
            }
        }

        public override Comparer<ICodeExplorerNode> SortComparer => CodeExplorerItemComparer.NodeType;

        public override bool Filtered => false;

        public override void Synchronize(List<Declaration> updated)
        {
            if (Declaration is null ||
                !(updated?.OfType<ProjectDeclaration>()
                    .FirstOrDefault(declaration => declaration.ProjectId.Equals(Declaration.ProjectId)) is ProjectDeclaration match))
            {
                Declaration = null;
                return;
            }

            Declaration = match;
            updated.Remove(Declaration);

            // Reference synchronization is deferred to AddNewChildren for 2 reasons. First, it doesn't make sense to sling around a List of
            // declaration for something that doesn't need it. Second (and more importantly), the priority can't be set without calling 
            // GetProjectReferenceModels, which hits the VBE COM interfaces. So, we only want to do that once. The bonus 3rd reason is that it
            // can be called from the ctor this way.

            SynchronizeChildren(updated.Where(declaration => declaration.ProjectId.Equals(Declaration.ProjectId)).ToList());
            updated.RemoveAll(declaration => declaration.ProjectId.Equals(Declaration.ProjectId));

            // Have to do this again - the project might have been saved or otherwise had the ProjectDisplayName changed.
            SetName();
        }

        protected sealed override void AddNewChildren(List<Declaration> updated)
        {
            if (updated is null)
            {
                return;
            }

            SynchronizeReferences();

            AddChildren(updated.GroupBy(declaration => declaration.RootFolder())
                .Where(folder => !string.IsNullOrEmpty(folder.Key))
                .Select(folder => new CodeExplorerCustomFolderViewModel(this, folder.Key, folder.Key, _vbe, folder)));
        }

        private void SynchronizeReferences()
        {
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
                AddChild(new CodeExplorerReferenceFolderViewModel(this, State.DeclarationFinder, type.ToList(), type.Key));
            }
        }

        private List<ReferenceModel> GetProjectReferenceModels()
        {
            var project = Declaration?.Project;
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
            _displayName = Declaration?.ProjectDisplayName ?? string.Empty;

            OnNameChanged();
        }
    }
}
