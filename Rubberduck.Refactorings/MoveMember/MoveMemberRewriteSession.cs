using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.Parsing.VBA.Parsing;
using Antlr4.Runtime;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberRewriteSession : IExecutableRewriteSession
    {
        /// <summary>
        /// The wrapped IExecutableRewriteSession
        /// </summary>
        IExecutableRewriteSession WrappedSession { get; }

        /// <summary>
        /// Caches declaration remove requests to support the use cases 
        /// for variables and constants in declaration lists
        /// and every variable/constant in the list is removed.
        /// All other Variables, Constants, and Members call rewriter.Remove 
        /// immediately
        /// </summary>
        void Remove(IEnumerable<Declaration> declarations);

        /// <summary>
        /// Caches declaration remove requests to support the use cases 
        /// for variables and constants  in declaration lists
        /// and every variable/constant in the list is removed.
        /// All other Variables, Constants, and Members call rewriter.Remove 
        /// immediately
        /// </summary>
        void Remove(Declaration target);
    }

    /// <summary>
    /// Extends IExecutableRewriteSession to intercept the
    /// IExecutableRewriteSession.TryRewrite() to insert a
    /// removal variables and constants IModuleRewriter.Remove requests
    /// </summary>
    public class MoveMemberRewriteSession : IMoveMemberRewriteSession
    {
        private IExecutableRewriteSession _rewriteSession;
        private Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>> RemovedVariables { set; get; } = new Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>>();
        private Dictionary<VBAParser.ConstStmtContext, HashSet<Declaration>> RemovedConstants { set; get; } = new Dictionary<VBAParser.ConstStmtContext, HashSet<Declaration>>();

        public MoveMemberRewriteSession(IExecutableRewriteSession rewriteSession)
        {
            _rewriteSession = rewriteSession;
        }

        public IExecutableRewriteSession WrappedSession => _rewriteSession;

        public IModuleRewriter CheckOutModuleRewriter(QualifiedModuleName qmn)
            => _rewriteSession.CheckOutModuleRewriter(qmn);

        public bool TryRewrite()
        {
            ExecuteCachedRemoveRequests();

            return _rewriteSession.TryRewrite();
        }

        public IReadOnlyCollection<QualifiedModuleName> CheckedOutModules => _rewriteSession.CheckedOutModules;

        public RewriteSessionState Status
        {
            get => _rewriteSession.Status;
            set { _rewriteSession.Status = value; }
        }

        public CodeKind TargetCodeKind => _rewriteSession.TargetCodeKind;

        public void Remove(IEnumerable<Declaration> declarations)
        {
            foreach (var declaration in declarations)
            {
                Remove(declaration);
            }
        }

        public void Remove(Declaration target)
        {
            if (target.DeclarationType.Equals(DeclarationType.Variable))
            {
                RemoveTarget<VBAParser.VariableListStmtContext, VBAParser.VariableSubStmtContext>(target, RemovedVariables);
                return;
            }
            if (target.DeclarationType.Equals(DeclarationType.Constant))
            {
                RemoveTarget<VBAParser.ConstStmtContext, VBAParser.ConstSubStmtContext>(target, RemovedConstants);
                return;
            }

            var rewriter = _rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);
            rewriter.Remove(target);
        }

        private void RemoveTarget<T, K>(Declaration target, Dictionary<T, HashSet<Declaration>> dictionary) where T : ParserRuleContext where K : ParserRuleContext
        {
            var declarationList = target.Context.GetAncestor<T>();

            if (declarationList is null || declarationList.children.Where(ch => ch is K).Count() == 1)
            {
                var rewriter = _rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);
                rewriter.Remove(target);
                return;
            }

            if (!dictionary.ContainsKey(declarationList))
            {
                dictionary.Add(declarationList, new HashSet<Declaration>());
            }
            dictionary[declarationList].Add(target);
        }

        private void ExecuteCachedRemoveRequests()
        {
            ExecuteCachedRemoveRequests<VBAParser.VariableListStmtContext, VBAParser.VariableSubStmtContext>(RemovedVariables);
            RemovedVariables = new Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>>();

            ExecuteCachedRemoveRequests<VBAParser.ConstStmtContext, VBAParser.ConstSubStmtContext>(RemovedConstants);
            RemovedConstants = new Dictionary<VBAParser.ConstStmtContext, HashSet<Declaration>>();
        }

        private void ExecuteCachedRemoveRequests<T, K>(Dictionary<T, HashSet<Declaration>> dictionary) where T: ParserRuleContext where K: ParserRuleContext
        {
            foreach (var key in dictionary.Keys)
            {
                if (dictionary[key].Count == 0)
                {
                    continue;
                }

                var rewriter = _rewriteSession.CheckOutModuleRewriter(dictionary[key].First().QualifiedModuleName);

                var toRemove = key.children.Where(ch => ch is K);
                if (toRemove.Count() == dictionary[key].Count)
                {
                    rewriter.Remove(key.Parent);
                }
                else
                {
                    foreach (var dec in dictionary[key])
                    {
                        rewriter.Remove(dec);
                    }
                }
            }
        }
    }
}
