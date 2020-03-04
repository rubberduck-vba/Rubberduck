using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberRefactoringStrategy
    {
        void RefactorRewrite(MoveMemberModel model, IRewriteSession rewriteSession, IRewritingManager rewritingManager, /*INewContentProvider movedContent,*/ bool asPreview = false);
        IMovedContentProvider NewDestinationModuleContent(MoveMemberModel model, IRewritingManager rewritingManager, IMovedContentProvider contentToMove);
        bool IsApplicable(MoveMemberModel model);
        bool IsExecutableModel(MoveMemberModel model, out string nonExecutableMessage);
    }

    public abstract class MoveMemberStrategyBase : IMoveMemberRefactoringStrategy
    {
        public abstract void RefactorRewrite(MoveMemberModel model, IRewriteSession rewriteSession, IRewritingManager rewritingManager, /*INewContentProvider movedContent,*/ bool asPreview = false);
        public abstract IMovedContentProvider NewDestinationModuleContent(MoveMemberModel model, IRewritingManager rewritingManager, IMovedContentProvider contentToMove);
        public abstract bool IsApplicable(MoveMemberModel model);
        public abstract bool IsExecutableModel(MoveMemberModel model, out string nonExecutableMessage);

        /// <summary>
        /// Clears entire VariableStmtContext or ConstantStmtContext
        /// when all the variables or constants declared in the list are removed.
        /// </summary>
        /// <param name="rewriteSession"></param>
        /// <param name="declarations"></param>
        protected static void RemoveDeclarations(IRewriteSession rewriteSession, IEnumerable<Declaration> declarations)
        {
            var removedVariables = new Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>>();
            var removedConstants = new Dictionary<VBAParser.ConstStmtContext, HashSet<Declaration>>();

            foreach (var declaration in declarations)
            {
                if (declaration.DeclarationType.Equals(DeclarationType.Variable))
                {
                    CacheListDeclaredElement<VBAParser.VariableListStmtContext, VBAParser.VariableSubStmtContext>(rewriteSession, declaration, removedVariables);
                    continue;
                }

                if (declaration.DeclarationType.Equals(DeclarationType.Constant))
                {
                    CacheListDeclaredElement<VBAParser.ConstStmtContext, VBAParser.ConstSubStmtContext>(rewriteSession, declaration, removedConstants);
                    continue;
                }

                var rewriter = rewriteSession.CheckOutModuleRewriter(declaration.QualifiedModuleName);
                rewriter.Remove(declaration);
            }

            ExecuteCachedRemoveRequests<VBAParser.VariableListStmtContext, VBAParser.VariableSubStmtContext>(rewriteSession, removedVariables);
            ExecuteCachedRemoveRequests<VBAParser.ConstStmtContext, VBAParser.ConstSubStmtContext>(rewriteSession, removedConstants);
        }

        private static void CacheListDeclaredElement<T, K>(IRewriteSession rewriteSession, Declaration target, Dictionary<T, HashSet<Declaration>> dictionary) where T : ParserRuleContext where K : ParserRuleContext
        {
            var declarationList = target.Context.GetAncestor<T>();

            if ((declarationList?.children.OfType<K>().Count() ?? 1) == 1)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);
                rewriter.Remove(target);
                return;
            }

            if (!dictionary.ContainsKey(declarationList))
            {
                dictionary.Add(declarationList, new HashSet<Declaration>());
            }
            dictionary[declarationList].Add(target);
        }

        private static void ExecuteCachedRemoveRequests<T, K>(IRewriteSession rewriteSession, Dictionary<T, HashSet<Declaration>> dictionary) where T : ParserRuleContext where K : ParserRuleContext
        {
            foreach (var key in dictionary.Keys.Where(k => dictionary[k].Any()))
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(dictionary[key].First().QualifiedModuleName);

                if (key.children.OfType<K>().Count() == dictionary[key].Count)
                {
                    rewriter.Remove(key.Parent);
                    continue;
                }

                foreach (var dec in dictionary[key])
                {
                    rewriter.Remove(dec);
                }
            }
        }
    }
}
