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
        private readonly IList<Node> _options = new List<Node>(); 
        private readonly IList<Node> _members = new List<Node>();

        public VBTreeListener(string project, string module)
        {
            _project = project;
            _module = module;
        }

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
            _options.Add(new OptionNode(GetSelection(context), _project, _module, OptionNode.VBOption.Explicit));
        }

        public override void EnterOptionBaseStmt(VisualBasic6Parser.OptionBaseStmtContext context)
        {
            _options.Add(new OptionNode(GetSelection(context), _project, _module, OptionNode.VBOption.Base, context.INTEGERLITERAL().Symbol.Text));
        }

        public override void EnterOptionCompareStmt(VisualBasic6Parser.OptionCompareStmtContext context)
        {
            _options.Add(new OptionNode(GetSelection(context), _project, _module, OptionNode.VBOption.Compare, context.children.Last().GetText()));
        }

        public override void ExitSubStmt(VisualBasic6Parser.SubStmtContext context)
        {
            var procedureName = context.ambiguousIdentifier().IDENTIFIER()[0].Symbol.Text;
            var accessibility = context.visibility().IsEmpty
                ? VBAccessibility.Public
                : (VBAccessibility)Enum.Parse(typeof(VBAccessibility), context.visibility().GetText());
            var node = new ProcedureNode(GetSelection(context), _project, _module, ProcedureNode.VBProcedureKind.Sub, procedureName, null, accessibility);
            var args = context.argList().arg().ToList();
        }

        public override void ExitFunctionStmt(VisualBasic6Parser.FunctionStmtContext context)
        {
            var procedureName = context.ambiguousIdentifier().IDENTIFIER()[0].Symbol.Text;
            var args = context.argList().arg().ToList();
            var returnType = context.asTypeClause().type();
        }

        public override void ExitPropertyGetStmt(VisualBasic6Parser.PropertyGetStmtContext context)
        {
            var propertyName = context.ambiguousIdentifier().IDENTIFIER()[0].Symbol.Text;
            var args = context.argList().arg().ToList();
            var returnType = context.asTypeClause().type();
        }

        public override void ExitPropertyLetStmt(VisualBasic6Parser.PropertyLetStmtContext context)
        {
            var propertyName = context.ambiguousIdentifier().IDENTIFIER()[0].Symbol.Text;
            var args = context.argList().arg().ToList();
        }

        public override void ExitPropertySetStmt(VisualBasic6Parser.PropertySetStmtContext context)
        {
            var propertyName = context.ambiguousIdentifier().IDENTIFIER()[0].Symbol.Text;
            var args = context.argList().arg().ToList();
        }
    }
}