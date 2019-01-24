using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Navigation.CodeExplorer
{
    public sealed class CodeExplorerReferenceViewModel : CodeExplorerItemViewModelBase
    {
        public CodeExplorerReferenceViewModel(ICodeExplorerNode parent, ReferenceModel reference) : base(parent, parent?.Declaration)
        {
            Reference = reference;
        }

        public ReferenceModel Reference { get; private set; }
        
        public override string Name => Reference?.Name ?? string.Empty;

        public override string NameWithSignature => Reference.Type == ReferenceKind.TypeLibrary
            ? $"{Name} ({Path.GetFileName(Reference.FullPath)} {Reference.Version})"
            : $"{Name} ({Path.GetFileName(Reference.FullPath)})";

        public override string PanelTitle => ToolTip;

        public override string Description => Reference?.FullPath ?? string.Empty;

        public override QualifiedSelection? QualifiedSelection => null;

        public override bool IsDimmed => !Reference?.IsUsed ?? true;

        public override bool IsErrorState
        {
            get => false;
            set { /* References can never be in an error state (in this context). */ }
        }

        public override string ToolTip => Reference?.Description ?? string.Empty;

        public int? Priority => Reference?.Priority;

        public bool Locked => Reference?.IsBuiltIn ?? false;

        public override Comparer<ICodeExplorerNode> SortComparer => CodeExplorerItemComparer.ReferencePriority;

        public void Synchronize(Declaration project, List<ReferenceModel> updated)
        {
            Declaration = project;

            var used = Reference?.IsUsed ?? false;
            Reference = updated.FirstOrDefault(reference => reference.Matches(Reference.ToReferenceInfo()));

            if (Reference == null)
            {
                return;
            }

            updated.Remove(Reference);

            if (used != Reference.IsUsed)
            {
                OnPropertyChanged(nameof(IsDimmed));
            }
        }
    }
}
