using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA.Grammar
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
        public static string GetLocalDeclarationSyntax(string keyword)
        {
            var keywords = new[] { ReservedKeywords.Dim, ReservedKeywords.Static };
            if (!keywords.Contains(keyword))
            {
                throw new InvalidOperationException("Keyword " + keyword + " is not valid in this context.");
            }

            return GetDeclarationSyntax(keyword);
        }

        /// <summary>
        /// Gets a regular expression pattern for matching a field declaration.
        /// </summary>
        /// <param name="keyword"><c>Private</c>, <c>Public</c>, or <c>Global</c>.</param>
        /// <returns></returns>
        public static string GetModuleDeclarationSyntax(string keyword)
        {
            var keywords = new[] { ReservedKeywords.Private, ReservedKeywords.Public, ReservedKeywords.Global };
            if (!keywords.Contains(keyword))
            {
                throw new InvalidOperationException("Keyword " + keyword + " is not valid in this context.");
            }

            return GetDeclarationSyntax(keyword);
        }

        private static string GetDeclarationSyntax(string keyword)
        {
            return "^" + keyword + @"(\s(?<identifier>[a-zA-Z][a-zA-Z0-9_]*)(?<specifier>[%&@!#$])?(?<array>\((?<size>(([0-9]+)\,?\s?)*|([0-9]+\sTo\s[0-9]+\,?\s?)+)\))?(?<as>\sAs(\s(?<initializer>New))?\s(?<reference>(((?<library>[a-zA-Z][a-zA-Z0-9_]*))\.)?(?<identifier>[a-zA-Z][a-zA-Z0-9_]*)))?(\,)?)+$";
        }

        /// <summary>
        /// Gets a regular expression pattern for matching a constant declaration.
        /// </summary>
        /// <remarks>
        /// Constants declared in class modules may only be <c>Private</c>.
        /// Constants declared at procedure scope cannot have an access modifier.
        /// </remarks>
        public static string GetConstantDeclarationSyntax()
        {
            return @"^((Private|Public|Global)\s)?Const\s(?<identifier>[a-zA-Z][a-zA-Z0-9_]*)(?<specifier>[%&@!#$])?(?<as>\sAs\s(?<reference>(((?<library>[a-zA-Z][a-zA-Z0-9_]*))\.)?(?<identifier>[a-zA-Z][a-zA-Z0-9_]*)))?\s\=\s(?<value>.*)$";
        }
    }
}
