using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    public static class ModuleBodyElementDeclarationExtensions
    {
        /// <summary>
        /// Returns ModuleBodyElementDeclaration signature with an ImprovedArgument list
        /// 1. Explicitly declares Property Let\Set value parameter as ByVal
        /// 2. Ensures UserDefined Type parameters are declared either explicitly or implicitly as ByRef
        /// </summary>
        /// <param name="declaration"></param>
        /// <returns></returns>
        public static string FullMemberSignature(this ModuleBodyElementDeclaration declaration,
                                    string accessibility = null,
                                    string newIdentifier = null)
        {
            var identifier = newIdentifier ?? declaration.IdentifierName;

            var fullSignatureFormat = string.Empty;
            switch (declaration.Context)
            {
                case VBAParser.SubStmtContext _:
                    fullSignatureFormat = $"{{0}} {Tokens.Sub} {identifier}({{1}}){{2}}";
                    break;
                case VBAParser.FunctionStmtContext _:
                    fullSignatureFormat = $"{{0}} {Tokens.Function} {identifier}({{1}}){{2}}";
                    break;
                case VBAParser.PropertyGetStmtContext _:
                    fullSignatureFormat = $"{{0}} {Tokens.Property} {Tokens.Get} {identifier}({{1}}){{2}}";
                    break;
                case VBAParser.PropertyLetStmtContext _:
                    fullSignatureFormat = $"{{0}} {Tokens.Property} {Tokens.Let} {identifier}({{1}}){{2}}";
                    break;
                case VBAParser.PropertySetStmtContext _:
                    fullSignatureFormat = $"{{0}} {Tokens.Property} {Tokens.Set} {identifier}({{1}}){{2}}";
                    break;
                default:
                    throw new ArgumentException();
            }

            var accessibilityToken = declaration.Accessibility.Equals(Accessibility.Implicit)
                ? Tokens.Public
                : $"{declaration.Accessibility.ToString()}";

            accessibilityToken = accessibility ?? accessibilityToken;

            var improvedArgList = ImprovedArgumentList(declaration);

            var asTypeSuffix = declaration.AsTypeName == null
                ? string.Empty
                : $" {Tokens.As} {declaration.AsTypeName}";

            var fullSignature = string.Format(fullSignatureFormat, accessibilityToken, improvedArgList, asTypeSuffix);
            return fullSignature.Trim();
        }

        public static string AsCodeBlock(this ModuleBodyElementDeclaration declaration,
                                            string content = null,
                                            string accessibility = null,
                                            string newIdentifier = null)
        {
            var endStatement = string.Empty;
            switch (declaration.Context)
            {
                case VBAParser.SubStmtContext _:
                    endStatement = $"{Tokens.End} {Tokens.Sub}";
                    break;
                case VBAParser.FunctionStmtContext _:
                    endStatement = $"{Tokens.End} {Tokens.Function}";
                    break;
                case VBAParser.PropertyGetStmtContext _:
                case VBAParser.PropertyLetStmtContext _:
                case VBAParser.PropertySetStmtContext _:
                    endStatement = $"{Tokens.End} {Tokens.Property}";
                    break;
                default:
                    throw new ArgumentException();
            }

            if (content != null)
            {
                return string.Format("{0}{1}{2}{1}{3}{1}",
                    FullMemberSignature(declaration, accessibility, newIdentifier),
                    Environment.NewLine,
                    content,
                    endStatement);
            }

            return string.Format("{0}{1}{2}{1}",
                FullMemberSignature(declaration, accessibility, newIdentifier),
                Environment.NewLine,
                endStatement);
        }

        /// <summary>
        /// 1. Explicitly declares Property Let\Set value parameter as ByVal
        /// 2. Ensures UserDefined Type parameters are declared either explicitly or implicitly as ByRef
        /// </summary>
        /// <param name="declaration"></param>
        /// <returns></returns>
        public static string ImprovedArgumentList(this ModuleBodyElementDeclaration declaration)
        {
            var arguments = Enumerable.Empty<string>();
            if (declaration is IParameterizedDeclaration parameterizedDeclaration)
            {
                arguments = parameterizedDeclaration.Parameters
                    .OrderBy(parameter => parameter.Selection)
                    .Select(parameter => BuildParameterDeclaration(
                            parameter,
                            parameter.Equals(parameterizedDeclaration.Parameters.LastOrDefault())
                                && declaration.DeclarationType.HasFlag(DeclarationType.Property)
                                && !declaration.DeclarationType.Equals(DeclarationType.PropertyGet)));
            }
            return $"{string.Join(", ", arguments)}";
        }

        private static string BuildParameterDeclaration(ParameterDeclaration parameter, bool forceExplicitByValAccess)
        {
            var accessibility = parameter.IsImplicitByRef
                ? string.Empty
                : parameter.IsByRef
                    ? Tokens.ByRef
                    : Tokens.ByVal;

            if (forceExplicitByValAccess)
            {
                accessibility = Tokens.ByVal;
            }

            if (accessibility.Equals(Tokens.ByVal)
                         && (parameter.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false))
            {
                accessibility = Tokens.ByRef;
            }

            var name = parameter.IsArray
                ? $"{parameter.IdentifierName}()"
                : parameter.IdentifierName;

            var optional = parameter.IsParamArray
               ? Tokens.ParamArray
               : parameter.IsOptional
                   ? Tokens.Optional
                   : string.Empty;

            var defaultValue = parameter.DefaultValue;

            return $"{FormatStandardElement(optional)}{FormatStandardElement(accessibility)}{FormatStandardElement(name)}{FormattedAsTypeName(parameter.AsTypeName)}{FormattedDefaultValue(defaultValue)}".Trim();
        }

        private static string FormatStandardElement(string element) => string.IsNullOrEmpty(element)
            ? string.Empty
            : $"{element} ";

        private static string FormattedAsTypeName(string AsTypeName) => string.IsNullOrEmpty(AsTypeName)
            ? string.Empty
            : $"As {AsTypeName} ";

        private static string FormattedDefaultValue(string DefaultValue) => string.IsNullOrEmpty(DefaultValue)
            ? string.Empty
            : $"= {DefaultValue}";
    }
}