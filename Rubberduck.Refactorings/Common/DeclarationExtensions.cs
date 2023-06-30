using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Common
{
    public static class DeclarationExtensions
    {
        public static bool IsVariable(this Declaration declaration)
            => declaration.DeclarationType.HasFlag(DeclarationType.Variable);

        public static bool IsMemberVariable(this Declaration declaration)
            => declaration.IsVariable() && !declaration.ParentDeclaration.IsMember();

        public static bool IsLocalVariable(this Declaration declaration)
            => declaration.IsVariable() && declaration.ParentDeclaration.IsMember();

        public static bool IsLocalConstant(this Declaration declaration)
            => declaration.IsConstant() && declaration.ParentDeclaration.IsMember();

        public static bool HasPrivateAccessibility(this Declaration declaration)
            => declaration.Accessibility.Equals(Accessibility.Private);

        public static bool IsMember(this Declaration declaration)
            => declaration.DeclarationType.HasFlag(DeclarationType.Member);

        public static bool IsConstant(this Declaration declaration)
            => declaration.DeclarationType.HasFlag(DeclarationType.Constant);

        public static bool IsUserDefinedTypeField(this Declaration declaration)
            => declaration.IsMemberVariable() && IsUserDefinedType(declaration);

        public static bool IsUserDefinedType(this Declaration declaration)
            => (declaration.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false);

        public static bool IsEnumField(this Declaration declaration)
            => declaration.IsMemberVariable() && (declaration.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.Enumeration) ?? false);

        public static bool IsDeclaredInList(this Declaration declaration)
        {
            return declaration.Context.TryGetAncestor<VBAParser.VariableListStmtContext>(out var varList)
                            && varList.ChildCount > 1;
        }

        public static bool IsMutatorProperty(this Declaration declaration)
            => declaration.DeclarationType.HasFlag(DeclarationType.PropertyLet)
            || declaration.DeclarationType.HasFlag(DeclarationType.PropertySet);
    }
}
