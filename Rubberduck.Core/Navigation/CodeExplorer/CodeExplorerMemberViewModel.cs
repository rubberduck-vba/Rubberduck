using System.Collections.Generic;
using System.Linq;
using System.Text;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Navigation.CodeExplorer
{
    public sealed class CodeExplorerMemberViewModel : CodeExplorerItemViewModel
    {
        public CodeExplorerMemberViewModel(ICodeExplorerNode parent, Declaration declaration, ref List<Declaration> declarations) : base(parent, declaration)
        {
            AddNewChildren(ref declarations);
            Name = DetermineMemberName(declaration);
        }

        public override string Name { get; }

        private string _signature;
        public override string NameWithSignature
        {
            get
            {
                if (_signature != null)
                {
                    return _signature;
                }

                if (Declaration is ValuedDeclaration value && !string.IsNullOrEmpty(value.Expression))
                {
                    _signature = $"{Name} = {value.Expression}";
                    return _signature;
                }

                if (!(Declaration.Context.children.FirstOrDefault(d => d is VBAParser.ArgListContext) is VBAParser.ArgListContext context))
                {
                    _signature = Name;
                }
                else if (Declaration is PropertyDeclaration)
                {
                    // 6 being the three-letter "get/let/set" + parens + space
                    _signature = Name.Insert(Name.Length - 6, RemoveExtraWhiteSpace(context.GetText()));
                }
                else
                {
                    _signature = Name + RemoveExtraWhiteSpace(context.GetText());
                }
                return _signature;
            }
        }

        public override bool IsObsolete =>
            Declaration.Annotations.Any(annotation => annotation is ObsoleteAnnotation);

        public static readonly DeclarationType[] SubMemberTypes =
        {
            DeclarationType.EnumerationMember,
            DeclarationType.UserDefinedTypeMember
        };

        public override void Synchronize(ref List<Declaration> updated)
        {
            base.Synchronize(ref updated);
            if (Declaration is null)
            {
                return;
            }

            // Parameter list might have changed - invalidate the signature.
            _signature = null;
            OnNameChanged();
        }

        protected override void AddNewChildren(ref List<Declaration> updated)
        {
            if (updated == null)
            {
                return;
            }

            var updates = updated.Where(item =>
                SubMemberTypes.Contains(item.DeclarationType) && item.ParentDeclaration.Equals(Declaration)).ToList();

            updated = updated.Except(updates.Concat(new[] { Declaration })).ToList();

            AddChildren(updates.Select(item => new CodeExplorerSubMemberViewModel(this, item)));
        }

        public override Comparer<ICodeExplorerNode> SortComparer
        {
            get
            {
                switch (SortOrder)
                {
                    case CodeExplorerSortOrder.Name:
                        return CodeExplorerItemComparer.Name;
                    case CodeExplorerSortOrder.CodeLine:
                        return CodeExplorerItemComparer.CodeLine;
                    case CodeExplorerSortOrder.DeclarationTypeThenName:
                        return CodeExplorerItemComparer.DeclarationTypeThenName;
                    case CodeExplorerSortOrder.DeclarationTypeThenCodeLine:
                        return CodeExplorerItemComparer.DeclarationTypeThenCodeLine;
                    default:
                        return CodeExplorerItemComparer.Name;
                }
            }
        }

        private static string RemoveExtraWhiteSpace(string value)
        {
            var newStr = new StringBuilder();
            var trimmedJoinedString = value.Replace(" _\r\n", " ").Trim();

            for (var i = 0; i < trimmedJoinedString.Length; i++)
            {
                // this will not throw because `Trim` ensures the first character is not whitespace
                if (char.IsWhiteSpace(trimmedJoinedString[i]) && char.IsWhiteSpace(trimmedJoinedString[i - 1]))
                {
                    continue;
                }

                newStr.Append(trimmedJoinedString[i]);
            }

            return newStr.ToString();
        }

        private static string DetermineMemberName(Declaration declaration)
        {
            switch (declaration.DeclarationType)
            {
                case DeclarationType.PropertyGet:
                    return $"{declaration.IdentifierName} ({Tokens.Get})";
                case DeclarationType.PropertyLet:
                    return $"{declaration.IdentifierName} ({Tokens.Let})";
                case DeclarationType.PropertySet:
                    return $"{declaration.IdentifierName} ({Tokens.Set})";
                case DeclarationType.Variable:
                    return declaration.IsArray ? $"{declaration.IdentifierName}()" : declaration.IdentifierName;
                default:
                    return declaration.IdentifierName;
            }
        }
    }
}
