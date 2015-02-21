using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class NodeBuildingListener : VBListenerBase
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

        public override void EnterSubStmt(VBParser.SubStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
        }

        public override void EnterFunctionStmt(VBParser.FunctionStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
        }

        public override void EnterPropertyGetStmt(VBParser.PropertyGetStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
        }

        public override void EnterPropertyLetStmt(VBParser.PropertyLetStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
        }

        public override void EnterPropertySetStmt(VBParser.PropertySetStmtContext context)
        {
            _currentNode = CreateProcedureNode(context);
        }

        private ProcedureNode CreateProcedureNode(dynamic context)
        {
            var procedureName = context.ambiguousIdentifier().GetText();
            var node = new ProcedureNode(context, _currentScope, procedureName);

            var args = context.argList().arg() as IReadOnlyList<VBParser.ArgContext>;
            if (args != null)
            {
                foreach (var arg in args)
                {
                    node.AddChild(new ParameterNode(arg, _currentScope));
                }
            }

            _currentScope = _project + "." + _module + "." + node.Name;
            return node;
        }

        public override void ExitOptionExplicitStmt(VBParser.OptionExplicitStmtContext context)
        {
            _members.Add(new OptionNode(context, _currentScope));
        }

        public override void ExitOptionBaseStmt(VBParser.OptionBaseStmtContext context)
        {
            _members.Add(new OptionNode(context, _currentScope));
        }

        public override void ExitOptionCompareStmt(VBParser.OptionCompareStmtContext context)
        {
            _members.Add(new OptionNode(context, _currentScope));
        }

        public override void ExitEnumerationStmt(VBParser.EnumerationStmtContext context)
        {
            _members.Add(new EnumNode(context, _currentScope));
        }

        public override void ExitSubStmt(VBParser.SubStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitFunctionStmt(VBParser.FunctionStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitPropertyGetStmt(VBParser.PropertyGetStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitPropertyLetStmt(VBParser.PropertyLetStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitPropertySetStmt(VBParser.PropertySetStmtContext context)
        {
            AddCurrentMember();
        }

        public override void ExitVariableStmt(VBParser.VariableStmtContext context)
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

        public override void ExitConstStmt(VBParser.ConstStmtContext context)
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

        public override void ExitTypeStmt(VBParser.TypeStmtContext context)
        {
            _members.Add(new TypeNode(context, _currentScope));
        }
    }
    
    public class VariableAssignmentListener : VariableUsageListener
    {
        public override void EnterVariableCallStmt(VBParser.VariableCallStmtContext context)
        {
            if (context.Parent.Parent.Parent is VBParser.LetStmtContext)
            {
                base.EnterVariableCallStmt(context);
            }
        }

        public VariableAssignmentListener(QualifiedModuleName qualifiedName) 
            : base(qualifiedName)
        {
        }
    }

    public class VariableUsageListener : VBListenerBase, IExtensionListener<VBParser.AmbiguousIdentifierContext>
    {
        private readonly IList<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _members = 
            new List<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();

        private readonly QualifiedModuleName _qualifiedName;

        public VariableUsageListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> Members { get { return _members; } }

        public override void EnterForNextStmt(VBParser.ForNextStmtContext context)
        {
            _members.Add(new QualifiedContext<VBParser.AmbiguousIdentifierContext>(_qualifiedName, context.AmbiguousIdentifier().First()));
        }

        public override void EnterVariableCallStmt(VBParser.VariableCallStmtContext context)
        {
            _members.Add(new QualifiedContext<VBParser.AmbiguousIdentifierContext>(_qualifiedName, context.AmbiguousIdentifier()));
        }

        public override void EnterWithStmt(VBParser.WithStmtContext context)
        {
            _members.Add(new QualifiedContext<VBParser.AmbiguousIdentifierContext>(_qualifiedName, context.ImplicitCallStmt_InStmt().ICS_S_VariableCall().VariableCallStmt().AmbiguousIdentifier()));
        }
    }
}