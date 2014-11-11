using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA.Grammar
{
    // todo: handle end-of-line comments
    // todo: handle multiple Const declarations in single instruction.

    internal static class VBAGrammar
    {
        /// <summary>
        /// Gets a regular expression pattern for matching a local variable declaration.
        /// </summary>
        /// <param name="keyword"><c>Dim</c> or <c>Static</c>.</param>
        /// <returns></returns>
        public static string LocalDeclarationSyntax(string keyword)
        {
            var keywords = new[] { ReservedKeywords.Dim, ReservedKeywords.Static };
            if (!keywords.Contains(keyword))
            {
                throw new InvalidOperationException("Keyword " + keyword + " is not valid in this context.");
            }

            return DeclarationSyntax(keyword);
        }

        /// <summary>
        /// Gets a regular expression pattern for matching a field declaration.
        /// </summary>
        /// <param name="keyword"><c>Private</c>, <c>Public</c>, or <c>Global</c>.</param>
        /// <returns></returns>
        public static string ModuleDeclarationSyntax(string keyword)
        {
            var keywords = new[] { ReservedKeywords.Private, ReservedKeywords.Public, ReservedKeywords.Global };
            if (!keywords.Contains(keyword))
            {
                throw new InvalidOperationException("Keyword " + keyword + " is not valid in this context.");
            }

            var pattern = DeclarationSyntax(keyword);
            return pattern;
        }

        private static string IdentifierSyntax { get { return @"(?<identifier>([a-zA-Z][a-zA-Z0-9_]*)|(\[[a-zA-Z0-9_]*\]))"; } }
        private static string ReferenceSyntax { get { return @"(?<reference>(((?<library>[a-zA-Z][a-zA-Z0-9_]*))\.)?)" + IdentifierSyntax; } }

        private static string DeclarationSyntax(string keyword)
        {
            return "^" + keyword + @"(\s" + IdentifierSyntax + @"(?<specifier>[%&@!#$])?(?<array>\((?<size>(([0-9]+)\,?\s?)*|([0-9]+\sTo\s[0-9]+\,?\s?)+)\))?(?<as>\sAs(\s(?<initializer>New))?\s" + ReferenceSyntax + @")?(\,)?)+$";
        }

        /// <summary>
        /// Gets a regular expression pattern for matching a constant declaration.
        /// </summary>
        /// <remarks>
        /// Constants declared in class modules may only be <c>Private</c>.
        /// Constants declared at procedure scope cannot have an access modifier.
        /// </remarks>
        public static string ConstantDeclarationSyntax()
        {
            return @"^((Private|Public|Global)\s)?Const\s" + IdentifierSyntax + @"(?<specifier>[%&@!#$])?(?<as>\sAs\s" + ReferenceSyntax + @")?\s\=\s(?<value>.*)$";
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
            return @"^((Private|Public)\s)?(?<ProcedureKind>(Sub|Function|Property (Get|Let|Set)))\s" + IdentifierSyntax;
        }
    }
}
