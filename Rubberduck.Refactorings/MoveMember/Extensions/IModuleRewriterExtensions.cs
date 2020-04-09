using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember.Extensions
{
    public static class IModuleRewriterExtensions
    {
        public static string GetText(this IModuleRewriter rewriter, Declaration declaration) 
            => rewriter.GetText(declaration.Context.Start.TokenIndex, declaration.Context.Stop.TokenIndex);

        public static void InsertAtEndOfFile(this IModuleRewriter rewriter, string content)
        {
            if (!string.IsNullOrEmpty(content))
            {
                rewriter.InsertBefore(rewriter.TokenStream.Size - 1, content);
            }
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
    }
}
