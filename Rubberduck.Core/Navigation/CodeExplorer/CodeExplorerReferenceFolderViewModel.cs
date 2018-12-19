using System.Windows.Media.Imaging;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerReferenceFolderViewModel : CodeExplorerItemViewModel
    {
        private readonly CodeExplorerProjectViewModel _parent;

        public CodeExplorerReferenceFolderViewModel(CodeExplorerProjectViewModel parent, ReferenceKind type = ReferenceKind.Project) : base(parent?.Declaration)
        {
            _parent = parent;
            ReferenceKind = type;

            CollapsedIcon = GetImageSource(Resources.CodeExplorer.CodeExplorerUI.ObjectAssembly);
            ExpandedIcon = GetImageSource(Resources.CodeExplorer.CodeExplorerUI.ObjectAssembly);

            AddReferenceNodes(type);
        }

        public ReferenceKind ReferenceKind { get; set; }

        public override string Name => ReferenceKind == ReferenceKind.TypeLibrary
            ? Resources.CodeExplorer.CodeExplorerUI.CodeExplorer_ProjectReferences
            : Resources.CodeExplorer.CodeExplorerUI.CodeExplorer_LibraryReferences;

        public override string NameWithSignature => Name;

        public override BitmapImage CollapsedIcon { get; }
        public override BitmapImage ExpandedIcon { get; }
        public override CodeExplorerItemViewModel Parent => _parent;
        public override QualifiedSelection? QualifiedSelection => null;

        private void AddReferenceNodes(ReferenceKind type)
        {
            var project = _parent?.Declaration?.Project;
            if (project == null)
            {
                return;
            }

            using (var references = project.References)
            {
                var priority = 1;
                foreach (var reference in references)
                {
                    if (reference.Type != type)
                    {
                        continue;
                    }

                    var model = new ReferenceModel(reference, priority++);
                    model.IsUsed = reference.IsBuiltIn ||
                                   _parent.State.DeclarationFinder.IsReferenceUsedInProject(
                                       _parent?.Declaration as ProjectDeclaration, model.ToReferenceInfo());

                    AddChild(new CodeExplorerReferenceViewModel(this, model));
                    reference.Dispose();
                }
            }
        }
    }
}
