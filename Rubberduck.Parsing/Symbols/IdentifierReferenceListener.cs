using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReferenceListener : VBAParserBaseListener
    {
        private readonly IdentifierReferenceResolver _resolver;

        public IdentifierReferenceListener(IdentifierReferenceResolver resolver)
        {
            _resolver = resolver;
            SetCurrentScope();
        }

        private void SetCurrentScope()
        {
            _resolver.SetCurrentScope();
        }

        private void SetCurrentScope(string identifier, DeclarationType type)
        {
            _resolver.SetCurrentScope(identifier, type);
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.Procedure);
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.Function);
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyGet);
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyLet);
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertySet);
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.Enumeration);
        }

        public override void ExitEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterTypeStmt(VBAParser.TypeStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.UserDefinedType);
        }

        public override void ExitTypeStmt(VBAParser.TypeStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterWithStmt(VBAParser.WithStmtContext context)
        {
            _resolver.EnterWithBlock(context);
        }

        public override void ExitWithStmt(VBAParser.WithStmtContext context)
        {
            _resolver.ExitWithBlock();
        }
        
        public override void EnterICS_B_ProcedureCall(VBAParser.ICS_B_ProcedureCallContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterICS_B_MemberProcedureCall(VBAParser.ICS_B_MemberProcedureCallContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterICS_S_VariableOrProcedureCall(VBAParser.ICS_S_VariableOrProcedureCallContext context)
        {
            if (context.Parent.GetType() != typeof(VBAParser.ICS_S_MemberCallContext))
            {
                _resolver.Resolve(context);
            }
        }

        public override void EnterICS_S_ProcedureOrArrayCall(VBAParser.ICS_S_ProcedureOrArrayCallContext context)
        {
            if (context.Parent.GetType() != typeof(VBAParser.ICS_S_MemberCallContext))
            {
                _resolver.Resolve(context);
            }
        }

        public override void EnterICS_S_MembersCall(VBAParser.ICS_S_MembersCallContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterICS_S_DictionaryCall(VBAParser.ICS_S_DictionaryCallContext context)
        {
            if (context.Parent.GetType() != typeof(VBAParser.ICS_S_MemberCallContext))
            {
                _resolver.Resolve(context);
            }
        }

        public override void EnterLetStmt(VBAParser.LetStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterSetStmt(VBAParser.SetStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterAsTypeClause(VBAParser.AsTypeClauseContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterForNextStmt(VBAParser.ForNextStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterForEachStmt(VBAParser.ForEachStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterImplementsStmt(VBAParser.ImplementsStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterRaiseEventStmt(VBAParser.RaiseEventStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterResumeStmt(VBAParser.ResumeStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterFieldLength(VBAParser.FieldLengthContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterVsAssign(VBAParser.VsAssignContext context)
        {
            _resolver.Resolve(context);
        }
    }
}