using System;
using System.Threading;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBA;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReferenceListener : VBABaseListener
    {
        private readonly IdentifierReferenceResolver _resolver;
        private readonly CancellationToken _token;
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

        private void TrySetCurrentScope(string identifier = null, DeclarationType? accessor = null)
        {
            try
            {
                if (identifier == null)
                {
                    _resolver.SetCurrentScope();
                }
                else
                {
                    _resolver.SetCurrentScope(identifier, accessor);
                }
            }
            catch (Exception exception)
            {
                // if we can't resolve the current scope, we can't resolve anything under it: force-cancel the walk.
                throw new WalkerCancelledException();
            }
        }

        public IdentifierReferenceListener(IdentifierReferenceResolver resolver, CancellationToken token)
        {
            _resolver = resolver;
            _token = token;
            TrySetCurrentScope();
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            TrySetCurrentScope(context.ambiguousIdentifier().GetText());
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            TrySetCurrentScope();
            OnMemberProcessed(context.ambiguousIdentifier().GetText());
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            TrySetCurrentScope(context.ambiguousIdentifier().GetText());
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            TrySetCurrentScope();
            OnMemberProcessed(context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            TrySetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyGet);
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            TrySetCurrentScope();
            OnMemberProcessed(context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            TrySetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyLet);
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            TrySetCurrentScope();
            OnMemberProcessed(context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            TrySetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertySet);
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            TrySetCurrentScope();
            OnMemberProcessed(context.ambiguousIdentifier().GetText());
        }

        public override void EnterWithStmt(VBAParser.WithStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.EnterWithBlock(context);
        }

        public override void ExitWithStmt(VBAParser.WithStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.ExitWithBlock();
        }
        
        public override void EnterICS_B_ProcedureCall(VBAParser.ICS_B_ProcedureCallContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterICS_B_MemberProcedureCall(VBAParser.ICS_B_MemberProcedureCallContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterICS_S_VariableOrProcedureCall(VBAParser.ICS_S_VariableOrProcedureCallContext context)
        {
            _token.ThrowIfCancellationRequested();
            if (context.Parent.GetType() != typeof(VBAParser.ICS_S_MemberCallContext))
            {
                _resolver.Resolve(context);
            }
        }

        public override void EnterICS_S_ProcedureOrArrayCall(VBAParser.ICS_S_ProcedureOrArrayCallContext context)
        {
            _token.ThrowIfCancellationRequested();
            if (context.Parent.GetType() != typeof(VBAParser.ICS_S_MemberCallContext))
            {
                _resolver.Resolve(context);
            }
        }

        public override void EnterICS_S_MembersCall(VBAParser.ICS_S_MembersCallContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterICS_S_DictionaryCall(VBAParser.ICS_S_DictionaryCallContext context)
        {
            _token.ThrowIfCancellationRequested();
            if (context.Parent.GetType() != typeof(VBAParser.ICS_S_MemberCallContext))
            {
                _resolver.Resolve(context);
            }
        }

        public override void EnterLetStmt(VBAParser.LetStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterSetStmt(VBAParser.SetStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterAsTypeClause(VBAParser.AsTypeClauseContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterForNextStmt(VBAParser.ForNextStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterForEachStmt(VBAParser.ForEachStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterImplementsStmt(VBAParser.ImplementsStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterRaiseEventStmt(VBAParser.RaiseEventStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterResumeStmt(VBAParser.ResumeStmtContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterFileNumber(VBAParser.FileNumberContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterArgDefaultValue(VBAParser.ArgDefaultValueContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterFieldLength(VBAParser.FieldLengthContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }

        public override void EnterVsAssign(VBAParser.VsAssignContext context)
        {
            _token.ThrowIfCancellationRequested();
            _resolver.Resolve(context);
        }
    }
}