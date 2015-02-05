using System;
using Antlr4.Runtime;
using Rubberduck.VBA.Grammar;

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
            : this(context, scope, localScope, VBProcedureKind.PropertyGet, context.visibility(), context.ambiguousIdentifier(), context.asTypeClause)
        {
        }

        public ProcedureNode(VisualBasic6Parser.FunctionStmtContext context, string scope, string localScope)
            : this(context, scope, localScope, VBProcedureKind.Function, context.visibility(), context.ambiguousIdentifier(), context.asTypeClause)
        {
        }

        public ProcedureNode(VisualBasic6Parser.SubStmtContext context, string scope, string localScope)
            : this(context, scope, localScope, VBProcedureKind.Sub, context.visibility(), context.ambiguousIdentifier(), null)
        {
        }

        private ProcedureNode(ParserRuleContext context, string scope, string localScope, 
                              VBProcedureKind kind, 
                              VisualBasic6Parser.VisibilityContext visibility, 
                              VisualBasic6Parser.AmbiguousIdentifierContext name, 
                              Func<VisualBasic6Parser.AsTypeClauseContext> asType)
            : base(context, scope, localScope)
        {
            _kind = kind;
            _name = name.GetText();
            _accessibility = visibility.GetAccessibility();

            if (asType != null)
            {
                var returnTypeClause = asType();
                _isImplicitReturnType = returnTypeClause == null;

                _returnType = returnTypeClause == null 
                                ? Tokens.Variant 
                                : returnTypeClause.type().GetText();
            }
        }

        private readonly string _name;
        public string Name { get { return _name; } }

        private readonly string _returnType;
        public string ReturnType { get { return _returnType; } }

        private readonly bool _isImplicitReturnType;
        public bool IsImplicitReturnType { get { return _isImplicitReturnType; } }

        private readonly VBProcedureKind _kind;
        public VBProcedureKind Kind { get { return _kind; } }

        private readonly VBAccessibility _accessibility;
        public VBAccessibility Accessibility { get { return _accessibility; } }
    }
}