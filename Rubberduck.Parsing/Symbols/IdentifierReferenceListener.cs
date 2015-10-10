using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReferenceListener : VBABaseListener
    {
        private readonly IdentifierReferenceResolver _resolver;
        public event EventHandler<MemberProcessedEventArgs> MemberProcessed;

        private void OnMemberProcessed(string name)
        {
            var handler = MemberProcessed;
            if (handler == null)
            {
                return;
            }

            var args = new MemberProcessedEventArgs(name);
            handler.Invoke(this, args);
        }

        public IdentifierReferenceListener(IdentifierReferenceResolver resolver)
        {
            _resolver = resolver;
            _resolver.SetCurrentScope();
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            _resolver.SetCurrentScope(context.ambiguousIdentifier().GetText());
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            _resolver.SetCurrentScope();
            OnMemberProcessed(context.ambiguousIdentifier().GetText());
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _resolver.SetCurrentScope(context.ambiguousIdentifier().GetText());
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _resolver.SetCurrentScope();
            OnMemberProcessed(context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _resolver.SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyGet);
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _resolver.SetCurrentScope();
            OnMemberProcessed(context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _resolver.SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyLet);
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _resolver.SetCurrentScope();
            OnMemberProcessed(context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _resolver.SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertySet);
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _resolver.SetCurrentScope();
            OnMemberProcessed(context.ambiguousIdentifier().GetText());
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
            if (context.Parent.GetType() != typeof (VBAParser.ICS_S_MemberCallContext))
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

        public override void EnterFileNumber(VBAParser.FileNumberContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterArgDefaultValue(VBAParser.ArgDefaultValueContext context)
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