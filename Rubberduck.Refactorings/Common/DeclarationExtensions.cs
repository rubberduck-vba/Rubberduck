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
        public static string FieldToPropertyBlock(this Declaration variable, DeclarationType letSetGetType, string propertyIdentifier, string accessibility, string content, string parameterIdentifier = "value")
        {
            var template = string.Join(Environment.NewLine, accessibility + " {0}{1} {2}{3}", $"{content}", Tokens.End + " {0}", string.Empty);

            var asType = variable.IsArray
                ? $"{Tokens.Variant}"
                : variable.IsEnumField() && variable.AsTypeDeclaration.HasPrivateAccessibility()
                        ? $"{Tokens.Long}"
                        : $"{variable.AsTypeName}";

            var paramAccessibility = variable.IsUserDefinedType() ? Tokens.ByRef : Tokens.ByVal;

            var letSetParameter = $"({paramAccessibility} {parameterIdentifier} {Tokens.As} {asType})";

            if (letSetGetType.Equals(DeclarationType.PropertyGet))
            {
                return string.Format(template, Tokens.Property, $" {Tokens.Get}", $"{propertyIdentifier}()", $" {Tokens.As} {asType}");
            }

            if (letSetGetType.Equals(DeclarationType.PropertyLet))
            {
                return string.Format(template, Tokens.Property, $" {Tokens.Let}", $"{propertyIdentifier}{letSetParameter}", string.Empty);
            }

            if (letSetGetType.Equals(DeclarationType.PropertySet))
            {
                return string.Format(template, Tokens.Property, $" {Tokens.Set}", $"{propertyIdentifier}{letSetParameter}", string.Empty);
            }

            return string.Empty;
        }
    }
}
