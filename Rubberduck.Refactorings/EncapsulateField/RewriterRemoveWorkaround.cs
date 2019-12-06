using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    //If all variables are removed from a list one by one the 
    //Accessiblity token is left behind
    public static class RewriterRemoveWorkAround
    {
        private static Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>> RemovedVariables { set; get; } = new Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>>();

        public static void Remove(Declaration target, IEncapsulateFieldRewriter rewriter)
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

        public static void RemoveFieldsDeclaredInLists(IEncapsulateFieldRewriter rewriter)
        {
            foreach (var key in RemovedVariables.Keys)
            {
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
