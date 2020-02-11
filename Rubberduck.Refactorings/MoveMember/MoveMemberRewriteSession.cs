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

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberRewriteSession : IExecutableRewriteSession
    {
        IExecutableRewriteSession WrappedSession { get; }
        void Remove(IEnumerable<Declaration> declarations);
        void Remove(Declaration target);
    }

    public class MoveMemberRewriteSession : IMoveMemberRewriteSession
    {
        private IExecutableRewriteSession _rewriteSession;
        private Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>> RemovedVariables { set; get; } = new Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>>();

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
            var varList = target.Context.GetAncestor<VBAParser.VariableListStmtContext>();

            if (varList is null || varList.children.Where(ch => ch is VBAParser.VariableSubStmtContext).Count() == 1)
            {
                var rewriter = _rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);
                rewriter.Remove(target);
                return;
            }

            if (!RemovedVariables.ContainsKey(varList))
            {
                RemovedVariables.Add(varList, new HashSet<Declaration>());
            }
            RemovedVariables[varList].Add(target);
        }

        private void ExecuteCachedRemoveRequests()
        {
            foreach (var key in RemovedVariables.Keys)
            {
                if (RemovedVariables[key].Count == 0)
                {
                    continue;
                }

                var rewriter = _rewriteSession.CheckOutModuleRewriter(RemovedVariables[key].First().QualifiedModuleName);

                var variables = key.children.Where(ch => ch is VBAParser.VariableSubStmtContext);
                if (variables.Count() == RemovedVariables[key].Count)
                {
                    rewriter.Remove(key.Parent);
                }
                else
                {
                    foreach (var dec in RemovedVariables[key])
                    {
                        rewriter.Remove(dec);
                    }
                }
            }
            RemovedVariables = new Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>>();
        }
    }
}
