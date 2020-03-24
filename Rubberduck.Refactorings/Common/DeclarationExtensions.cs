using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using System;
using System.Collections.Generic;
using System.Linq;

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

        /// <summary>
        /// Generates a Property Member code block specified by the letSetGet DeclarationType argument.
        /// </summary>
        /// <param name="variable"></param>
        /// <param name="letSetGetType"></param>
        /// <param name="propertyIdentifier"></param>
        /// <param name="accessibility"></param>
        /// <param name="content"></param>
        /// <param name="parameterIdentifier"></param>
        /// <returns></returns>
        public static string FieldToPropertyBlock(this Declaration variable, DeclarationType letSetGetType, string propertyIdentifier, string accessibility = null, string content = null, string parameterIdentifier = null)
        {
            //"value" is the default
            var propertyValueParam = parameterIdentifier ?? Resources.RubberduckUI.EncapsulateField_DefaultPropertyParameter;

            var propertyEndStmt = $"{Tokens.End} {Tokens.Property}";

            var asType = variable.IsArray
                ? $"{Tokens.Variant}"
                : variable.IsEnumField() && variable.AsTypeDeclaration.HasPrivateAccessibility()
                        ? $"{Tokens.Long}"
                        : $"{variable.AsTypeName}";

            var asTypeClause = $"{Tokens.As} {asType}";

            var paramAccessibility = variable.IsUserDefinedType() ? Tokens.ByRef : Tokens.ByVal;

            var letSetParameter = $"{paramAccessibility} {propertyValueParam} {Tokens.As} {asType}";

            switch (letSetGetType)
            {
                case DeclarationType.PropertyGet:
                    return string.Join(Environment.NewLine, $"{accessibility ?? Tokens.Public} {PropertyTypeStatement(letSetGetType)} {propertyIdentifier}() {asTypeClause}", content, propertyEndStmt);
                case DeclarationType.PropertyLet:
                case DeclarationType.PropertySet:
                    return string.Join(Environment.NewLine, $"{accessibility ?? Tokens.Public} {PropertyTypeStatement(letSetGetType)} {propertyIdentifier}({letSetParameter})", content, propertyEndStmt);
                default:
                    throw new ArgumentException();
            }
        }

        private static string PropertyTypeStatement(DeclarationType declarationType)
        {
            switch (declarationType)
            {
                case DeclarationType.PropertyGet:
                    return $"{Tokens.Property} {Tokens.Get}";
                case DeclarationType.PropertyLet:
                    return $"{Tokens.Property} {Tokens.Let}";
                case DeclarationType.PropertySet:
                    return $"{Tokens.Property} {Tokens.Set}";
                default:
                    throw new ArgumentException();
            }

        }
    }
}
