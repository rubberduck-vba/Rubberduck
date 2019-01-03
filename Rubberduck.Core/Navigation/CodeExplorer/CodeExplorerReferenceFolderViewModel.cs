using System.Windows.Media.Imaging;
using Rubberduck.AddRemoveReferences;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerReferenceFolderViewModel : CodeExplorerItemViewModel
    {
        private readonly CodeExplorerProjectViewModel _parent;

        public CodeExplorerReferenceFolderViewModel(CodeExplorerProjectViewModel parent)
        {
            _parent = parent;
            CollapsedIcon = GetImageSource(Resources.CodeExplorer.CodeExplorerUI.ObjectAssembly);
            ExpandedIcon = GetImageSource(Resources.CodeExplorer.CodeExplorerUI.ObjectAssembly);
            AddReferenceNodes();
        }

        public override string Name => "References";
        public override string NameWithSignature => "References";
        public override BitmapImage CollapsedIcon { get; }
        public override BitmapImage ExpandedIcon { get; }
        public override CodeExplorerItemViewModel Parent => _parent;
        public override QualifiedSelection? QualifiedSelection => null;

        private void AddReferenceNodes()
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
                    AddChild(new CodeExplorerReferenceViewModel(this, new ReferenceModel(reference, priority++)));
                    reference.Dispose();
                }
            }
        }
    }
}
