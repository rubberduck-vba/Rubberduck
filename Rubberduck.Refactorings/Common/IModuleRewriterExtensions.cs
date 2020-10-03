using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public static class IModuleRewriterExtensions
    {
        /// <summary>
        /// Removes variable declaration and subsequent <c>VBAParser.EndOfStatementContext</c>
        /// depending on the <paramref name="removeEndOfStmtContext"/> flag.
        /// This function is intended to be called only once per rewriter within a given <c>ModuleRewriteSession</c>.  
        /// </summary>
        /// <remarks>
        /// Calling this function with <paramref name="removeEndOfStmtContext"/> defaulted to <c>true</c>
        /// avoids leaving residual newlines between the deleted declaration and the next declaration. 
        /// The one-time call constraint is required for scenarios where variables to delete are declared in a list.  Specifically,
        /// the use case where all the variables in the list are to be removed.
        /// If the variables to remove are not declared in a list, then this function can be called multiple times.
        /// </remarks>
        public static void RemoveVariables(this IModuleRewriter rewriter, IEnumerable<VariableDeclaration> toRemove, bool removeEndOfStmtContext = true)
        {
            if (!toRemove.Any())
            {
                return;
            }

            var fieldsToDeleteByListContext = toRemove.Distinct()
                .ToLookup(f => f.Context.GetAncestor<VBAParser.VariableListStmtContext>());

            foreach (var fieldsToDelete in fieldsToDeleteByListContext)
            {
                var variableList = fieldsToDelete.Key.children.OfType<VBAParser.VariableSubStmtContext>();

                if (variableList.Count() == fieldsToDelete.Count())
                {
                    if (fieldsToDelete.First().ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
                    {
                        rewriter.RemoveDeclarationContext<VBAParser.ModuleDeclarationsElementContext>(fieldsToDelete.First(), removeEndOfStmtContext);
                    }
                    else
                    {
                        rewriter.RemoveDeclarationContext<VBAParser.BlockStmtContext>(fieldsToDelete.First(), removeEndOfStmtContext);
                    }
                    continue;
                }

                foreach (var target in fieldsToDelete)
                {
                    rewriter.Remove(target);
                }
            }
        }

        /// <summary>
        /// Removes a member declaration and subsequent <c>VBAParser.EndOfStatementContext</c>
        /// depending on the <paramref name="removeEndOfStmtContext"/> flag.
        /// </summary>
        /// <remarks>
        /// Calling this function with <paramref name="removeEndOfStmtContext"/> defaulted to <c>true</c>
        /// avoids leaving residual newlines between the deleted declaration and the next declaration. 
        /// </remarks>
        public static void RemoveMember(this IModuleRewriter rewriter, ModuleBodyElementDeclaration target, bool removeEndOfStmtContext = true)
        {
            RemoveMembers(rewriter, new ModuleBodyElementDeclaration[] { target }, removeEndOfStmtContext);
        }

        /// <summary>
        /// Removes member declarations and subsequent <c>VBAParser.EndOfStatementContext</c>
        /// depending on the <paramref name="removeEndOfStmtContext"/> flag.
        /// </summary>
        /// <remarks>
        /// Calling this function with <paramref name="removeEndOfStmtContext"/> defaulted to <c>true</c>
        /// avoids leaving residual newlines between the deleted declaration and the next declaration. 
        /// </remarks>
        public static void RemoveMembers(this IModuleRewriter rewriter, IEnumerable<ModuleBodyElementDeclaration> toRemove, bool removeEndOfStmtContext = true)
        {
            if (!toRemove.Any())
            {
                return;
            }

            foreach (var member in toRemove)
            {
                rewriter.RemoveDeclarationContext<VBAParser.ModuleBodyElementContext>(member, removeEndOfStmtContext);
            }
        }

        private static void RemoveDeclarationContext<T>(this IModuleRewriter rewriter, Declaration declaration, bool removeEndOfStmtContext = true) where T : ParserRuleContext
        {
            if (!declaration.Context.TryGetAncestor<T>(out var elementContext))
            {
                throw new ArgumentException();
            }

            rewriter.Remove(elementContext);
            if (removeEndOfStmtContext && elementContext.TryGetFollowingContext<VBAParser.EndOfStatementContext>(out var nextContext))
            {
                rewriter.Remove(nextContext);
            }
        }
    }
}
