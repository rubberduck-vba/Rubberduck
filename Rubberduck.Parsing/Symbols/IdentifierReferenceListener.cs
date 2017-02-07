using Antlr4.Runtime.Misc;
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
            SetCurrentScope(Identifier.GetName(context.subroutineName().identifier()), DeclarationType.Procedure);
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SetCurrentScope(Identifier.GetName(context.functionName().identifier()), DeclarationType.Function);
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SetCurrentScope(Identifier.GetName(context.functionName().identifier()), DeclarationType.PropertyGet);
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope(Identifier.GetName(context.subroutineName().identifier()), DeclarationType.PropertyLet);
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SetCurrentScope(Identifier.GetName(context.subroutineName().identifier()), DeclarationType.PropertySet);
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            SetCurrentScope(Identifier.GetName(context.identifier()), DeclarationType.Enumeration);
            _resolver.Resolve(context);
        }

        public override void ExitEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPublicTypeDeclaration(VBAParser.PublicTypeDeclarationContext context)
        {
            SetCurrentScope(Identifier.GetName(context.udtDeclaration().untypedIdentifier()), DeclarationType.UserDefinedType);
        }

        public override void ExitPublicTypeDeclaration(VBAParser.PublicTypeDeclarationContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPrivateTypeDeclaration(VBAParser.PrivateTypeDeclarationContext context)
        {
            SetCurrentScope(Identifier.GetName(context.udtDeclaration().untypedIdentifier()), DeclarationType.UserDefinedType);
        }

        public override void ExitPrivateTypeDeclaration(VBAParser.PrivateTypeDeclarationContext context)
        {
            SetCurrentScope();
        }

        public override void EnterArrayDim(VBAParser.ArrayDimContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterWithStmt(VBAParser.WithStmtContext context)
        {
            _resolver.EnterWithBlock(context);
        }

        public override void ExitWithStmt(VBAParser.WithStmtContext context)
        {
            _resolver.ExitWithBlock();
        }

        public override void EnterIfStmt(VBAParser.IfStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterSingleLineIfStmt(VBAParser.SingleLineIfStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterSelectCaseStmt(VBAParser.SelectCaseStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterGoToStmt(VBAParser.GoToStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterOnGoToStmt(VBAParser.OnGoToStmtContext context)
        {       
            _resolver.Resolve(context);
        }

        public override void EnterGoSubStmt([NotNull] VBAParser.GoSubStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterOnGoSubStmt(VBAParser.OnGoSubStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterCallStmt(VBAParser.CallStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterConstStmt(VBAParser.ConstStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterRedimStmt(VBAParser.RedimStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterNameStmt(VBAParser.NameStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterEraseStmt(VBAParser.EraseStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterLsetStmt(VBAParser.LsetStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterRsetStmt(VBAParser.RsetStmtContext context)
        {
            _resolver.Resolve(context);
        }
        
        public override void EnterWhileWendStmt(VBAParser.WhileWendStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterOpenStmt(VBAParser.OpenStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterCloseStmt(VBAParser.CloseStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterSeekStmt([NotNull] VBAParser.SeekStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterLockStmt([NotNull] VBAParser.LockStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterUnlockStmt([NotNull] VBAParser.UnlockStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterLineInputStmt([NotNull] VBAParser.LineInputStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterWidthStmt([NotNull] VBAParser.WidthStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterPrintStmt([NotNull] VBAParser.PrintStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterWriteStmt([NotNull] VBAParser.WriteStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterInputStmt([NotNull] VBAParser.InputStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterPutStmt([NotNull] VBAParser.PutStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterGetStmt([NotNull] VBAParser.GetStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterOnErrorStmt(VBAParser.OnErrorStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterErrorStmt(VBAParser.ErrorStmtContext context)
        {
            _resolver.Resolve(context);
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

        public override void EnterDoLoopStmt([NotNull] VBAParser.DoLoopStmtContext context)
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

        public override void EnterLineSpecialForm(VBAParser.LineSpecialFormContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterCircleSpecialForm(VBAParser.CircleSpecialFormContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterScaleSpecialForm(VBAParser.ScaleSpecialFormContext context)
        {
            _resolver.Resolve(context);
        }

        public override void EnterDebugPrintStmt([NotNull] VBAParser.DebugPrintStmtContext context)
        {
            _resolver.Resolve(context);
        }
    }
}
