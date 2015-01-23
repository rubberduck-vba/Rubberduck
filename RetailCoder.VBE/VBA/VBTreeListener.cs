using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Extensions;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.VBA
{
    public static class ParserRuleContextExtensions
    {
        public static Selection GetSelection(this ParserRuleContext context)
        {
            // adding +1 because ANTLR indexes are 0-based, but VBE's are 1-based.
            return new Selection(
                context.Start.Line + 1,
                context.Start.StartIndex + 1, // todo: figure out why this is off and how to fix it
                context.Stop.Line + 1,
                context.Stop.StopIndex + 1); // todo: figure out why this is off and how to fix it
        }
    }

    public partial class VBTreeListener : VisualBasic6BaseListener
    {
        private readonly string _project;
        private readonly string _module;
        private readonly IList<Node> _members = new List<Node>();

        private string _currentScope;
        private Node _currentNode;

        public VBTreeListener(string project, string module)
        {
            _project = project;
            _module = module;
            _currentScope = project + "." + module;
        }

        public Node Root
        {
            get { return new ModuleNode(null, _project, _module, _members); }
        }

        private void AddCurrentMember()
        {
            _members.Add(_currentNode);
            _currentNode = null;
        }

        public override void EnterSubStmt(VisualBasic6Parser.SubStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
            _currentScope = _project + "." + _module + "." + ((ProcedureNode) _currentNode).Name;
        }

        public override void EnterFunctionStmt(VisualBasic6Parser.FunctionStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
            _currentScope = _project + "." + _module + "." + ((ProcedureNode)_currentNode).Name;
        }

        public override void EnterPropertyGetStmt(VisualBasic6Parser.PropertyGetStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
            _currentScope = _project + "." + _module + "." + ((ProcedureNode)_currentNode).Name;
        }

        public override void EnterPropertyLetStmt(VisualBasic6Parser.PropertyLetStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
            _currentScope = _project + "." + _module + "." + ((ProcedureNode)_currentNode).Name;
        }

        public override void EnterPropertySetStmt(VisualBasic6Parser.PropertySetStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
            _currentScope = _project + "." + _module + "." + ((ProcedureNode)_currentNode).Name;
        }

        private ProcedureNode CreateProcedureNode(dynamic context)
        {
            var procedureName = context.ambiguousIdentifier().IDENTIFIER()[0].Symbol.Text;
            var node = new ProcedureNode(context, _currentScope, procedureName);
            
            var args = context.argList().arg() as IReadOnlyList<VisualBasic6Parser.ArgContext>;
            if (args != null)
            {
                foreach (var arg in args)
                {
                    node.AddChild(new ParameterNode(arg, _currentScope));
                }
            }

            return node;
        }

        public override void ExitOptionExplicitStmt(VisualBasic6Parser.OptionExplicitStmtContext context)
        {
            _members.Add(new OptionNode(context, _currentScope));
        }

        public override void ExitOptionBaseStmt(VisualBasic6Parser.OptionBaseStmtContext context)
        {
            _members.Add(new OptionNode(context, _currentScope));
        }

        public override void ExitOptionCompareStmt(VisualBasic6Parser.OptionCompareStmtContext context)
        {
            _members.Add(new OptionNode(context, _currentScope));
        }

        public override void ExitEnumerationStmt(VisualBasic6Parser.EnumerationStmtContext context)
        {
            _members.Add(new EnumNode(context, _currentScope));
        }

        public override void ExitSubStmt(VisualBasic6Parser.SubStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitFunctionStmt(VisualBasic6Parser.FunctionStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitPropertyGetStmt(VisualBasic6Parser.PropertyGetStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitPropertyLetStmt(VisualBasic6Parser.PropertyLetStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitPropertySetStmt(VisualBasic6Parser.PropertySetStmtContext context)
        {
            AddCurrentMember();
        }
 
    }
}