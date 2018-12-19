using System;
using System.IO;
using System.Windows.Media.Imaging;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigation.CodeExplorer
{
    public sealed class CodeExplorerReferenceViewModel : CodeExplorerItemViewModel
    {
        public CodeExplorerReferenceViewModel(CodeExplorerReferenceFolderViewModel parent, ReferenceModel reference) : base(parent?.Declaration)
        {
            Parent = parent;
            Reference = reference;
            IsDimmed = !Reference.IsUsed;
        }

        public override string NameWithSignature => $"{Reference.Name} ({Path.GetFileName(Reference.FullPath)} {Reference.Version})";

        public override string Name => Reference.Name;

        public override string ToolTip => Reference.Description;

        public override CodeExplorerItemViewModel Parent { get; }

        public override QualifiedSelection? QualifiedSelection => null;

        public override BitmapImage CollapsedIcon => GetIcon();

        public override BitmapImage ExpandedIcon => GetIcon();

        public ReferenceModel Reference { get; }

        public int? Priority => Reference.Priority;

        public bool Locked => Reference.IsBuiltIn;

        private BitmapImage GetIcon()
        {
            if (Reference.Status.HasFlag(ReferenceStatus.Broken))
            {
                GetImageSource(CodeExplorerUI.BrokenReference);
            }

            return Reference.IsBuiltIn ? GetImageSource(CodeExplorerUI.LockedReference) : GetImageSource(CodeExplorerUI.Reference);
        }
    }
}
