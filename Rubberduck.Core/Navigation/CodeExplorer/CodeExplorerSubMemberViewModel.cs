using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Navigation.CodeExplorer
{
    public sealed class CodeExplorerSubMemberViewModel : CodeExplorerItemViewModel
    {
        public static readonly DeclarationType[] SubMemberTypes =
        {
            DeclarationType.EnumerationMember,
            DeclarationType.UserDefinedTypeMember
        };

        private readonly string _signature = string.Empty;

        public CodeExplorerSubMemberViewModel(ICodeExplorerNode parent, Declaration declaration) : base(parent, declaration)
        {
            if (Declaration is ValuedDeclaration value && !string.IsNullOrEmpty(value.Expression))
            {
                _signature = $" = {value.Expression}";
            }
        }

        public override string Name => Declaration?.IdentifierName ?? string.Empty;

        public override string NameWithSignature => $"{Name}{_signature}";

        public override void Synchronize(ref List<Declaration> updated)
        {
            var signature = _signature;

            base.Synchronize(ref updated);
            if (Declaration is null || _signature.Equals(signature))
            {
                return;
            }

            // Signature changed - update the UI.
            OnNameChanged();
        }

        public override Comparer<ICodeExplorerNode> SortComparer =>
            SortOrder.HasFlag(CodeExplorerSortOrder.Name)
                ? CodeExplorerItemComparer.Name
                : CodeExplorerItemComparer.CodeLine;

        // Bottom level node. This is a NOP.
        protected override void AddNewChildren(ref List<Declaration> updated) { }
    }
}
