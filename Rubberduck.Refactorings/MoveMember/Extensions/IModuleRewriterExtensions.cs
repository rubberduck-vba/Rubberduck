using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember.Extensions
{
    public static class IModuleRewriterExtensions
    {
        public static string GetText(this IModuleRewriter rewriter, Declaration declaration)
        {
            return rewriter.GetText(declaration.Context.Start.TokenIndex, declaration.Context.Stop.TokenIndex);
        }

        public static string GetText(this IModuleRewriter rewriter, int maxConsecutiveNewLines)
        {
            var result = rewriter.GetText();
            var target = string.Join(string.Empty, Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines).ToList());
            var replacement = string.Join(string.Empty, Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines - 1).ToList());
            for (var counter = 1; counter < 10 && result.Contains(target); counter++)
            {
                result = result.Replace(target, replacement);
            }
            return result;
        }

        public static void InsertAtEndOfFile(this IModuleRewriter rewriter, string content)
        {
            if (content == string.Empty)
            {
                return;
            }
            rewriter.InsertBefore(rewriter.TokenStream.Size - 1, content);
        }

        public static void SetMemberAccessibility(this IModuleRewriter rewriter, Declaration element, string accessibility)
        {
            var visibilityContext = element.Context.GetChild<VBAParser.VisibilityContext>();
            if (visibilityContext != null)
            {
                rewriter.Replace(visibilityContext, accessibility);
            }
            else if (element.IsMember())
            {
                rewriter.InsertBefore(element.Context.Start.TokenIndex, $"{accessibility} ");
            }
        }

        public static void RemoveMemberAccess(this IModuleRewriter rewriter, IEnumerable<IdentifierReference> memberReferences)
        {
            foreach (var idRef in memberReferences)
            {
                if (idRef.Context.Parent is VBAParser.MemberAccessExprContext maec)
                {
                    Debug.Assert(maec.ChildCount == 3, "MemberAccessExprContext child contexts does not equal 3");
                    rewriter.Replace(maec, maec.children[2].GetText());
                }
            }
        }

        public static void RemoveWithMemberAccess(this IModuleRewriter rewriter, IEnumerable<IdentifierReference> references)
        {
            foreach (var withMemberAccessExprContext in references.Where(rf => rf.Context.Parent is VBAParser.WithMemberAccessExprContext).Select(rf => rf.Context.Parent as VBAParser.WithMemberAccessExprContext))
            {
                rewriter.RemoveRange(withMemberAccessExprContext.Start.TokenIndex, withMemberAccessExprContext.Start.TokenIndex);
            }
        }

        public static void Rename(this IModuleRewriter rewriter, Declaration target, string newName)
        {
            if (!(target.Context is IIdentifierContext context))
            {
                throw new ArgumentException();
            }

            rewriter.Replace(context.IdentifierTokens, newName);
        }

        public static void Rename(this IModuleRewriter rewriter, IdentifierReference idRef, string newName)
        {
            //if (!(idRef.Context is IIdentifierContext context))
            //{
            //    throw new ArgumentException();
            //}

            rewriter.Replace(idRef.Context, newName);
        }

        public static void Rename(this IModuleRewriter rewriter, IIdentifierContext identifierContext, string newName) 
            => rewriter.Replace(identifierContext.IdentifierTokens, newName);
    }

}
