using System;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;

namespace Rubberduck.VBA.Nodes
{
    public class ProcedureNode : Node
    {
        public enum VBProcedureKind
        {
            Sub,
            Function,
            PropertyGet,
            PropertyLet,
            PropertySet
        }

        public ProcedureNode(VisualBasic6Parser.PropertySetStmtContext context, string scope, string localScope)
            : this(context, scope, localScope, VBProcedureKind.PropertySet, context.visibility(), context.ambiguousIdentifier(), null)
        {
        }

        public ProcedureNode(VisualBasic6Parser.PropertyLetStmtContext context, string scope, string localScope)
            : this(context, scope, localScope, VBProcedureKind.PropertyLet, context.visibility(), context.ambiguousIdentifier(), null)
        {
        }

        public ProcedureNode(VisualBasic6Parser.PropertyGetStmtContext context, string scope, string localScope)
            : this(context, scope, localScope, VBProcedureKind.PropertyGet, context.visibility(), context.ambiguousIdentifier(), context.asTypeClause())
        {
        }

        public ProcedureNode(VisualBasic6Parser.FunctionStmtContext context, string scope, string localScope)
            : this(context, scope, localScope, VBProcedureKind.Function, context.visibility(), context.ambiguousIdentifier(), context.asTypeClause())
        {
        }

        public ProcedureNode(VisualBasic6Parser.SubStmtContext context, string scope, string localScope)
            : this(context, scope, localScope, VBProcedureKind.Sub, context.visibility(), context.ambiguousIdentifier(), null)
        {
        }

        private ProcedureNode(ParserRuleContext context, string scope, string localScope, 
                              VBProcedureKind kind, 
                              IParseTree visibility, 
                              VisualBasic6Parser.AmbiguousIdentifierContext name, 
                              VisualBasic6Parser.AsTypeClauseContext asType)
            : base(context, scope, localScope)
        {
            _kind = kind;
            _name = name.GetText();
            if (visibility == null || string.IsNullOrEmpty(visibility.GetText()))
            {
                _accessibility = VBAccessibility.Implicit;
            }
            else
            {
                _accessibility = (VBAccessibility) Enum.Parse(typeof (VBAccessibility), visibility.GetText());
            }

            if (asType != null)
            {
                _returnType = asType.type().GetText();
            }
        }

        private readonly string _name;
        public string Name { get { return _name; } }

        private readonly string _returnType;
        public string ReturnType { get { return _returnType; } }

        private readonly VBProcedureKind _kind;
        public VBProcedureKind Kind { get { return _kind; } }

        private readonly VBAccessibility _accessibility;
        public VBAccessibility Accessibility { get { return _accessibility; } }
    }
}