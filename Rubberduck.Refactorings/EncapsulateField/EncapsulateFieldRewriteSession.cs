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

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldRewriteSession
    {
        IExecutableRewriteSession RewriteSession { get; }
        IModuleRewriter CheckOutModuleRewriter(QualifiedModuleName qmn);
        bool TryRewrite();
        string CreatePreview(QualifiedModuleName qmn);
        void Remove(Declaration target, IModuleRewriter rewriter);
    }

    public class EncapsulateFieldRewriteSession : IEncapsulateFieldRewriteSession
    {
        private IExecutableRewriteSession _rewriteSession;
        private Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>> RemovedVariables { set; get; } = new Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>>();

        public EncapsulateFieldRewriteSession(IExecutableRewriteSession rewriteSession)
        {
            _rewriteSession = rewriteSession;
        }

        public IExecutableRewriteSession RewriteSession => _rewriteSession;

        public IModuleRewriter CheckOutModuleRewriter(QualifiedModuleName qmn) 
            => _rewriteSession.CheckOutModuleRewriter(qmn);

        public bool TryRewrite()
        {
            ExecuteCachedRemoveRequests();

            return _rewriteSession.TryRewrite();
        }

        public string CreatePreview(QualifiedModuleName qmn)
        {
            ExecuteCachedRemoveRequests();

            var previewRewriter = _rewriteSession.CheckOutModuleRewriter(qmn);

            return previewRewriter.GetText(maxConsecutiveNewLines: 3);
        }

        public void Remove(Declaration target, IModuleRewriter rewriter)
        {
            var varList = target.Context.GetAncestor<VBAParser.VariableListStmtContext>();
            if (varList.children.Where(ch => ch is VBAParser.VariableSubStmtContext).Count() == 1)
            {
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

                var rewriter = RewriteSession.CheckOutModuleRewriter(RemovedVariables[key].First().QualifiedModuleName);

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
