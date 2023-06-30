using System;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Adjusts a Sub procedure to be a Function procedure, and updates all usages.
    /// </summary>
    /// <inspections>
    /// <inspection name="ProcedureCanBeWrittenAsFunctionInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="false" project="false" all="false" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     GetValue value
    ///     Debug.Print value
    /// End Sub
    /// 
    /// Private Sub GetValue(ByRef value As Long)
    ///     value = 42
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     value = GetValue(value)
    ///     Debug.Print value
    /// End Sub
    /// 
    /// Private Function GetValue(ByVal value As Long) As Long
    ///     value = 42
    ///     GetValue = value
    /// End Function
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class ChangeProcedureToFunctionQuickFix : QuickFixBase
    {
        public ChangeProcedureToFunctionQuickFix()
            : base(typeof(ProcedureCanBeWrittenAsFunctionInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var parameterizedDeclaration = (IParameterizedDeclaration) result.Target;
            var arg = parameterizedDeclaration.Parameters.First(p => p.IsByRef || p.IsImplicitByRef);
            var argIndex = parameterizedDeclaration.Parameters.IndexOf(arg);
            
            UpdateSignature(result.Target, arg, rewriteSession);
            foreach (var reference in result.Target.References.Where(reference => !reference.IsDefaultMemberAccess))
            {
                UpdateCall(reference, argIndex, rewriteSession);
            }
        }

        private void UpdateSignature(Declaration target, ParameterDeclaration arg, IRewriteSession rewriteSession)
        {
            var subStmt = (VBAParser.SubStmtContext) target.Context;
            var argContext = (VBAParser.ArgContext)arg.Context;

            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);

            rewriter.Replace(subStmt.SUB(), Tokens.Function);
            rewriter.Replace(subStmt.END_SUB(), "End Function");

            rewriter.InsertAfter(subStmt.argList().Stop.TokenIndex, $" As {arg.AsTypeName}");

            if (arg.IsByRef)
            {
                rewriter.Replace(argContext.BYREF(), Tokens.ByVal);
            }
            else if (arg.IsImplicitByRef)
            {
                rewriter.InsertBefore(argContext.unrestrictedIdentifier().Start.TokenIndex, Tokens.ByVal);
            }

            var returnStmt = $"    {subStmt.subroutineName().GetText()} = {argContext.unrestrictedIdentifier().GetText()}{Environment.NewLine}";
            rewriter.InsertBefore(subStmt.END_SUB().Symbol.TokenIndex, returnStmt);
        }

        private void UpdateCall(IdentifierReference reference, int argIndex, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(reference.QualifiedModuleName);
            var callStmtContext = reference.Context.GetAncestor<VBAParser.CallStmtContext>();
            var argListContext = callStmtContext.GetChild<VBAParser.ArgumentListContext>();

            var arg = argListContext.argument()[argIndex];
            var argName = arg.positionalArgument()?.argumentExpression() ?? arg.namedArgument().argumentExpression();

            rewriter.InsertBefore(callStmtContext.Start.TokenIndex, $"{argName.GetText()} = ");
            rewriter.Replace(callStmtContext.whiteSpace(), "(");
            rewriter.InsertAfter(argListContext.Stop.TokenIndex, ")");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ProcedureShouldBeFunctionInspectionQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
        public override bool CanFixAll => false;
    }
}