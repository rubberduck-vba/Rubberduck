using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ChangeProcedureToFunctionQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type> { typeof(ProcedureCanBeWrittenAsFunctionInspection) };

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        private readonly RubberduckParserState _state;

        public ChangeProcedureToFunctionQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public void Fix(IInspectionResult result)
        {
            var parameterizedDeclaration = (IParameterizedDeclaration) result.Target;
            var arg = parameterizedDeclaration.Parameters.Cast<ParameterDeclaration>().First(p => p.IsByRef || p.IsImplicitByRef);
            var argIndex = parameterizedDeclaration.Parameters.ToList().IndexOf(arg);
            
            UpdateSignature(result.Target, arg);
            foreach (var reference in result.Target.References)
            {
                UpdateCall(reference, argIndex);
            }
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.ProcedureShouldBeFunctionInspectionQuickFix;
        }

        private void UpdateSignature(Declaration target, ParameterDeclaration arg)
        {
            var subStmt = (VBAParser.SubStmtContext) target.Context;
            var argContext = (VBAParser.ArgContext)arg.Context;

            var rewriter = _state.GetRewriter(target);

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

            var returnStmt = "    " + subStmt.subroutineName().GetText() + " = " + argContext.unrestrictedIdentifier().GetText() + Environment.NewLine;
            rewriter.InsertBefore(subStmt.END_SUB().Symbol.TokenIndex, returnStmt);
        }

        private void UpdateCall(IdentifierReference reference, int argIndex)
        {
            var rewriter = _state.GetRewriter(reference.QualifiedModuleName);
            var callStmtContext = ParserRuleContextHelper.GetParent<VBAParser.CallStmtContext>(reference.Context);
            var argListContext = ParserRuleContextHelper.GetChild<VBAParser.ArgumentListContext>(callStmtContext);

            var arg = argListContext.argument()[argIndex];
            var argName = arg.positionalArgument()?.argumentExpression() ?? arg.namedArgument().argumentExpression();

            rewriter.InsertBefore(callStmtContext.Start.TokenIndex, $"{argName.GetText()} = ");
            rewriter.Replace(callStmtContext.whiteSpace(), "(");
            rewriter.InsertAfter(argListContext.Stop.TokenIndex, ")");
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => false;
        public bool CanFixInProject => false;
    }
}