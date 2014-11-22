using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Parser.Grammar
{
    [ComVisible(false)]
    public static class VBAGrammar
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
                               .ToList();
        }

        public static string IdentifierDeclarationSyntax
        {
            get
            {
                return "(?<declarations>(?:" + IdentifierSyntax +
                       @"(?<specifier>[%&@!#$])?(?<array>\((?<size>(([0-9]+)\,?\s?)*|([0-9]+\sTo\s[0-9]+\,?\s?)+)\))?(?<as>\sAs(\s(?<initializer>New))?\s" +
                       ReferenceSyntax + @")?)(?:\,\s)?)+";
            }
        }

        public static string DeclarationKeywordsSyntax
        {
            get
            {
                return
                    @"^(?:(?:(?<keywords>(?:(?:(?<accessibility>Private|Public|Global)\s)|(?<accessibility>Private|Public|Global)\s)?(?:(?<keyword>Private|Public|Friend|Global|Dim|Const|Static|Sub|Function|Property\sGet|Property\sLet|Property\sSet|Enum|Type|Declare\sFunction)))\s)?)";
            }
        }

        public static string GeneralDeclarationSyntax
        {
            get { return DeclarationKeywordsSyntax + "(?<expression>.*)?"; }
        }

        public static string LabelSyntax
        {
            get { return @"^(?<identifier>[a-zA-Z][a-zA-Z0-9_]*)\:$"; }
        }

        public static string EnumSyntax
        {
            get { return @"^((Private|Public|Global)\s)?Enum\s" + IdentifierSyntax; }
        }


        public static string EnumMemberSyntax
        {
            get { return @"^" + IdentifierSyntax + @"(\s\=\s(?<value>.*))?$"; }
        }

        public static string UserDefinedTypeSyntax
        {
            get { return @"^((Private|Public|Global)\s)?Type\s" + IdentifierSyntax; }
        }

        public static string ProcedureSyntax
        {
            get
            {
                return @"^(?<accessibility>(Friend|Private|Public)\s)?(?:(?<kind>Sub|Function|Property\s(Get|Let|Set)))\s" +
                       IdentifierSyntax +
                       @"\((?<parameters>.*)\)(?:\sAs\s(?<reference>(((?<library>[a-zA-Z][a-zA-Z0-9_]*))\.)?(?<identifier>([a-zA-Z][a-zA-Z0-9_]*)|\[[a-zA-Z0-9_]*\])))?";
            }
        }

        public static string ParameterSyntax
        {
            get
            {
                return @"(?:(?:(?:\s?(?<optional>Optional)\s)?(?<by>ByRef|ByVal|ParamArray)?\s))?(?:" + IdentifierSyntax +
                       @"(?<specifier>[%&@!#$])?(?<array>\((?<size>(?:(?:[0-9]+)\,?\s?)*|(?:[0-9]+\sTo\s[0-9]+\,?\s?)+)\))?(?<as>\sAs(?:\s" +
                       ReferenceSyntax + @")?))";
            }
        }

        public static string IfBlockSyntax
        {
            get { return @"If\s(?<condition>.*)\sThen(?:\s(?<expression>.*))?"; }
        }

        public static string ForLoopSyntax
        {
            get { return @"For\s" + IdentifierSyntax + @"\s=\s(?<lower>.*)\sTo\s(?<upper>.*)(?:\sStep\s(?<step>.*))?"; }
        }

        public static string ForEachLoopSyntax
        {
            get { return @"For\sEach\s" + IdentifierSyntax + @"\sIn\s" + ReferenceSyntax; }
        }

        public static string TypeConversionSyntax
        {
            get { return @"(?<keyword>CBool|CByte|CCur|CDate|CDbl|CInt|CLng|CSng|CStr|CVar)\((?<expression>.*)\)"; }
        }
    }
}
