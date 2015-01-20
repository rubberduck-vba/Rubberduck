using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Extensions;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.VBA
{
    public class VBTreeListener : VisualBasic6BaseListener
    {
        private readonly string _project;
        private readonly string _module;
        private readonly IList<Node> _members = new List<Node>();

        private string _currentScope;
        private Node _currentNode;
        private IList<Node> _children;

        public VBTreeListener(string project, string module)
        {
            _project = project;
            _module = module;
            _currentScope = project + "." + module;
        }

        public Node Root { get { return new ModuleNode(Selection.Empty, _project, _module, _members); } }

        private Selection GetSelection(ParserRuleContext context)
        {
            return new Selection(
                context.Start.Line + 1, 
                context.Start.StartIndex + 1, 
                context.Stop.Line + 1, 
                context.Stop.StopIndex + 1);
        }

        public override void EnterOptionExplicitStmt(VisualBasic6Parser.OptionExplicitStmtContext context)
        {
            _members.Add(new OptionNode(GetSelection(context), _project, _module, OptionNode.VBOption.Explicit));
        }

        public override void EnterOptionBaseStmt(VisualBasic6Parser.OptionBaseStmtContext context)
        {
            _members.Add(new OptionNode(GetSelection(context), _project, _module, OptionNode.VBOption.Base, context.INTEGERLITERAL().Symbol.Text));
        }

        public override void EnterOptionCompareStmt(VisualBasic6Parser.OptionCompareStmtContext context)
        {
            _members.Add(new OptionNode(GetSelection(context), _project, _module, OptionNode.VBOption.Compare, context.children.Last().GetText()));
        }

        public override void EnterSubStmt(VisualBasic6Parser.SubStmtContext context)
        {
            _currentNode = GetProcedureNode(context, ProcedureNode.VBProcedureKind.Sub);
            _children = new List<Node>();
            _currentScope = _project + "." + _module + "." + ((ProcedureNode) _currentNode).Name;
        }

        public override void EnterFunctionStmt(VisualBasic6Parser.FunctionStmtContext context)
        {
            _currentNode = GetProcedureNode(context, ProcedureNode.VBProcedureKind.Function);
            _children = new List<Node>();
            _currentScope = _project + "." + _module + "." + ((ProcedureNode)_currentNode).Name;
        }

        public override void EnterPropertyGetStmt(VisualBasic6Parser.PropertyGetStmtContext context)
        {
            _currentNode = GetProcedureNode(context, ProcedureNode.VBProcedureKind.PropertyGet);
            _children = new List<Node>();
            _currentScope = _project + "." + _module + "." + ((ProcedureNode)_currentNode).Name;
        }

        public override void EnterPropertyLetStmt(VisualBasic6Parser.PropertyLetStmtContext context)
        {
            _currentNode = GetProcedureNode(context, ProcedureNode.VBProcedureKind.PropertyLet);
            _children = new List<Node>();
            _currentScope = _project + "." + _module + "." + ((ProcedureNode)_currentNode).Name;
        }

        public override void EnterPropertySetStmt(VisualBasic6Parser.PropertySetStmtContext context)
        {
            _currentNode = GetProcedureNode(context, ProcedureNode.VBProcedureKind.PropertySet);
            _children = new List<Node>();
            _currentScope = _project + "." + _module + "." + ((ProcedureNode)_currentNode).Name;
        }

        private ProcedureNode GetProcedureNode(dynamic context, ProcedureNode.VBProcedureKind kind, string returnType = null)
        {
            var procedureName = context.ambiguousIdentifier().IDENTIFIER()[0].Symbol.Text;
            var accessibility = context.visibility().IsEmpty
                ? VBAccessibility.Public
                : (VBAccessibility)Enum.Parse(typeof(VBAccessibility), context.visibility().GetText());

            var node = new ProcedureNode(GetSelection(context), _project, _module, kind, procedureName, returnType, accessibility);
            var args = context.argList().arg() as IReadOnlyList<VisualBasic6Parser.ArgContext>;
            if (args != null)
            {
                foreach (var arg in args)
                {
                    ParameterNode.VBParameterType parameterType;
                    var byVal = arg.BYVAL();
                    var byRef = arg.BYREF();
                    if (byVal == null && byRef == null)
                    {
                        parameterType = ParameterNode.VBParameterType.ImplicitByRef;
                    }
                    else
                    {
                        parameterType = byRef == null
                            ? ParameterNode.VBParameterType.ByVal
                            : ParameterNode.VBParameterType.ByRef;
                    }

                    var name = arg.ambiguousIdentifier().GetText();
                    var type = arg.asTypeClause().type().GetText();

                    var isOptional = arg.OPTIONAL() != null;

                    var param = new ParameterNode(GetSelection(arg), _project, _module, parameterType, name, type, isOptional);
                    node.Children.Add(param);
                }
            }

            foreach (var child in _children)
            {
                node.Children.Add(child);
            }

            return node;
        }

        private void AddCurrentMember()
        {
            _members.Add(_currentNode);
            _currentNode = null;
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