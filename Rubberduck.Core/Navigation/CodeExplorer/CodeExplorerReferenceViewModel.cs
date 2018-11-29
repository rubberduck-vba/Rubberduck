using System;
using System.Windows.Media.Imaging;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerReferenceViewModel : CodeExplorerItemViewModel
    {
        private ReferenceModel _reference;

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

        public int Priority => _reference.Priority;
        public bool Locked => _reference.IsBuiltIn;

        private BitmapImage GetIcon()
        {
            switch (_reference.Status)
            {
                case ReferenceStatus.None:
                case ReferenceStatus.Loaded:
                    return _reference.IsBuiltIn ? GetImageSource(CodeExplorerUI.LockedReference) : GetImageSource(CodeExplorerUI.Reference);
                case ReferenceStatus.BuiltIn:
                    return GetImageSource(CodeExplorerUI.ObjectLibrary);
                case ReferenceStatus.Broken:
                case ReferenceStatus.Removed:
                    return GetImageSource(CodeExplorerUI.BrokenReference);
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
    }
}
