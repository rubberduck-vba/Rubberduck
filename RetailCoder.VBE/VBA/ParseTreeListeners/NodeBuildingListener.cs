using System.Collections.Generic;
using Rubberduck.Parsing;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class NodeBuildingListener : VBABaseListener
    {
        private readonly string _project;
        private readonly string _module;
        private readonly IList<Node> _members = new List<Node>();

        private string _currentScope;
        private Node _currentNode;

        public NodeBuildingListener(string project, string module)
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

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
        }

        private ProcedureNode CreateProcedureNode(dynamic context)
        {
            var procedureName = context.ambiguousIdentifier().GetText();
            var node = new ProcedureNode(context, _currentScope, procedureName);

            var args = context.argList().arg() as IReadOnlyList<VBAParser.ArgContext>;
            if (args != null)
            {
                foreach (VBAParser.ArgContext arg in args)
                {
                    node.AddChild(new ParameterNode(arg, _currentScope));
                }
            }

            _currentScope = _project + "." + _module + "." + node.Name;
            return node;
        }

        public override void ExitOptionExplicitStmt(VBAParser.OptionExplicitStmtContext context)
        {
            _members.Add(new OptionNode(context, _currentScope));
        }

        public override void ExitOptionBaseStmt(VBAParser.OptionBaseStmtContext context)
        {
            _members.Add(new OptionNode(context, _currentScope));
        }

        public override void ExitOptionCompareStmt(VBAParser.OptionCompareStmtContext context)
        {
            _members.Add(new OptionNode(context, _currentScope));
        }

        public override void ExitEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            _members.Add(new EnumNode(context, _currentScope));
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitVariableStmt(VBAParser.VariableStmtContext context)
        {
            var node = new VariableDeclarationNode(context, _currentScope);
            if (_currentNode == null)
            {
                _members.Add(node);
            }
            else
            {
                _currentNode.AddChild(node);
            }
        }

        public override void ExitConstStmt(VBAParser.ConstStmtContext context)
        {
            var node = new ConstDeclarationNode(context, _currentScope);
            if (_currentNode == null)
            {
                _members.Add(node);
            }
            else
            {
                _currentNode.AddChild(node);
            }
        }

        public override void ExitTypeStmt(VBAParser.TypeStmtContext context)
        {
            _members.Add(new TypeNode(context, _currentScope));
        }
    }
}