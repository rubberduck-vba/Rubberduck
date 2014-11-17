using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Rubberduck.VBA.Parser.Grammar
{
    // todo: handle end-of-line comments
    // todo: handle multiple Const declarations in single instruction.

    internal static class VBAGrammar
    {
        private static string IdentifierSyntax { get { return @"(?<identifier>(?:[a-zA-Z][a-zA-Z0-9_]*)|(?:\[[a-zA-Z0-9_]*\]))"; } }
        private static string ReferenceSyntax { get { return @"(?<reference>(?:(?:(?<library>[a-zA-Z][a-zA-Z0-9_]*))\.)?" + IdentifierSyntax + ")"; } }

        /// <summary>
        /// Finds all implementations of <see cref="SyntaxBase"/> in the Rubberduck assembly.
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<ISyntax> GetGrammarSyntax()
        {
            return Assembly.GetExecutingAssembly()
                               .GetTypes()
                               .Where(type => type.BaseType == typeof(SyntaxBase))
                               .Select(type =>
                               {
                                   var constructorInfo = type.GetConstructor(Type.EmptyTypes);
                                   return constructorInfo != null ? constructorInfo.Invoke(Type.EmptyTypes) : null;
                               })
                               .Cast<ISyntax>()
                               .Where(syntax => !syntax.IsChildNodeSyntax)
                               .ToList();
        }

        public static string IdentifierDeclarationSyntax()
        {
            return "(?<declarations>(?:" + IdentifierSyntax + @"(?<specifier>[%&@!#$])?(?<array>\((?<size>(([0-9]+)\,?\s?)*|([0-9]+\sTo\s[0-9]+\,?\s?)+)\))?(?<as>\sAs(\s(?<initializer>New))?\s" + ReferenceSyntax + @")?)(?:\,\s)?)+";
        }

        public static string DeclarationKeywordsSyntax()
        {
            return @"^(?:(?:(?<keywords>(?:(?:(?<accessibility>Private|Public|Global)\s)|(?<accessibility>Private|Public|Global)\s)?(?:(?<keyword>Private|Public|Friend|Global|Dim|Const|Static|Sub|Function|Property\sGet|Property\sLet|Property\sSet|Enum|Type|Declare\sFunction)))\s)?)";
        }

        public static string GeneralDeclarationSyntax()
        {
            return DeclarationKeywordsSyntax() + "(?<expression>.*)?";
        }

        public static string LabelSyntax()
        {
            return @"^(?<identifier>[a-zA-Z][a-zA-Z0-9_]*)\:$";
        }

        public static string EnumSyntax()
        {
            return @"^((Private|Public|Global)\s)?Enum\s" + IdentifierSyntax;
        }

        public static string EnumMemberSyntax()
        {
            return @"^" + IdentifierSyntax + @"(\s\=\s(?<value>.*))?$";
        }

        public static string ProcedureSyntax()
        {
            return @"^(?<accessibility>(Friend|Private|Public)\s)?(?:(?<kind>Sub|Function|Property\s(Get|Let|Set)))\s" + IdentifierSyntax + @"\((?<parameters>.*)\)(?:\sAs\s(?<reference>(((?<library>[a-zA-Z][a-zA-Z0-9_]*))\.)?(?<identifier>([a-zA-Z][a-zA-Z0-9_]*)|\[[a-zA-Z0-9_]*\])))?";
        }

        public static string ParameterSyntax()
        {
            return @"(?:(?:(?:\s?(?<optional>Optional)\s)?(?<by>ByRef|ByVal|ParamArray)?\s))?(?:" + IdentifierSyntax + @"(?<specifier>[%&@!#$])?(?<array>\((?<size>(?:(?:[0-9]+)\,?\s?)*|(?:[0-9]+\sTo\s[0-9]+\,?\s?)+)\))?(?<as>\sAs(?:\s" + ReferenceSyntax + @")?))";
        }
    }
}
