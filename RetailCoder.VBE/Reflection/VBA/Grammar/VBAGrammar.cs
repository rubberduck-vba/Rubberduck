using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA.Grammar
{
    // todo: handle end-of-line comments
    // todo: handle multiple Const declarations in single instruction.

    internal static class VBAGrammar
    {
        private static string IdentifierSyntax { get { return @"(?<identifier>([a-zA-Z][a-zA-Z0-9_]*)|(\[[a-zA-Z0-9_]*\]))"; } }
        private static string ReferenceSyntax { get { return @"(?<reference>(((?<library>[a-zA-Z][a-zA-Z0-9_]*))\.)?" + IdentifierSyntax + ")"; } }

        public static string IdentifierDeclarationSyntax()
        {
            return "(" + IdentifierSyntax + @"(?<specifier>[%&@!#$])?(?<array>\((?<size>(([0-9]+)\,?\s?)*|([0-9]+\sTo\s[0-9]+\,?\s?)+)\))?(?<as>\sAs(\s(?<initializer>New))?\s" + ReferenceSyntax + @")?(\,)?)+$";
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
            return @"^((Private|Public)\s)?(?:(?<ProcedureKind>Sub|Function|Property)\s(Get|Let|Set))\s" + IdentifierSyntax;
        }
    }
}
