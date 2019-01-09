using System;
using System.Windows.Media.Imaging;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerReferenceViewModel : CodeExplorerItemViewModel
    {
        private readonly ReferenceModel _reference;

        public CodeExplorerReferenceViewModel(CodeExplorerReferenceFolderViewModel parent, ReferenceModel reference)
        {
            Parent = parent;
            _reference = reference;
        }

        public override string NameWithSignature => $"{_reference.Name} ({_reference.Version})";
        public override string Name => _reference.Description + Environment.NewLine + _reference.FullPath;
        public override CodeExplorerItemViewModel Parent { get; }
        public override QualifiedSelection? QualifiedSelection => null;

        public override BitmapImage CollapsedIcon => GetIcon();
        public override BitmapImage ExpandedIcon => GetIcon();

        public int? Priority => _reference.Priority;
        public bool Locked => _reference.IsBuiltIn;

        private BitmapImage GetIcon()
        {
            if (_reference.Status.HasFlag(ReferenceStatus.Broken))
            {
                GetImageSource(CodeExplorerUI.BrokenReference);
            }

            return _reference.IsBuiltIn ? GetImageSource(CodeExplorerUI.LockedReference) : GetImageSource(CodeExplorerUI.Reference);
        }
    }
}
