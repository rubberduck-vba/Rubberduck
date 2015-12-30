using System;
using System.Collections.Generic;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ProcedureShouldBeFunctionInspectionResult : CodeInspectionResultBase
    {
       private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

       public ProcedureShouldBeFunctionInspectionResult(IInspection inspection, QualifiedContext<VBAParser.ArgListContext> argListQualifiedContext, QualifiedContext<VBAParser.SubStmtContext> subStmtQualifiedContext, QualifiedContext<VBAParser.ArgContext> argQualifiedContext)
           : base(inspection,
                string.Format(inspection.Description, subStmtQualifiedContext.Context.ambiguousIdentifier().GetText()),
                subStmtQualifiedContext.ModuleName,
                subStmtQualifiedContext.Context.ambiguousIdentifier())
        {
            _quickFixes = new[]
            {
                new ChangeProcedureToFunction(argListQualifiedContext, subStmtQualifiedContext, argQualifiedContext, QualifiedSelection), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class ChangeProcedureToFunction : CodeInspectionQuickFix
    {
        private readonly QualifiedContext<VBAParser.ArgListContext> _argListQualifiedContext;
        private readonly QualifiedContext<VBAParser.SubStmtContext> _subStmtQualifiedContext;
        private readonly QualifiedContext<VBAParser.ArgContext> _argQualifiedContext;

        public ChangeProcedureToFunction(QualifiedContext<VBAParser.ArgListContext> argListQualifiedContext,
                                         QualifiedContext<VBAParser.SubStmtContext> subStmtQualifiedContext,
                                         QualifiedContext<VBAParser.ArgContext> argQualifiedContext,
                                         QualifiedSelection selection)
            : base(subStmtQualifiedContext.Context, selection, InspectionsUI.ProcedureShouldBeFunctionInspectionQuickFix)
        {
            _argListQualifiedContext = argListQualifiedContext;
            _subStmtQualifiedContext = subStmtQualifiedContext;
            _argQualifiedContext = argQualifiedContext;
        }

        public override void Fix()
        {
            var argListText = _argListQualifiedContext.Context.GetText();
            var subStmtText = _subStmtQualifiedContext.Context.GetText();
            var argText = _argQualifiedContext.Context.GetText();

            var newFunctionWithoutReturn = subStmtText.Insert(subStmtText.IndexOf(argListText, StringComparison.Ordinal) + argListText.Length,
                                                              _argQualifiedContext.Context.asTypeClause().GetText())
                                                      .Replace("Sub", "Function")
                                                      .Replace(argText,
                                                               argText.Contains("ByRef ")
                                                                 ? argText.Replace("ByRef ", "ByVal ")
                                                                 : "ByVal " + argText);

            var newfunctionWithReturn = newFunctionWithoutReturn
                .Insert(newFunctionWithoutReturn.LastIndexOf(Environment.NewLine, StringComparison.Ordinal),
                    "    " + _subStmtQualifiedContext.Context.ambiguousIdentifier().GetText() + " = " +
                    _argQualifiedContext.Context.ambiguousIdentifier().GetText());

            var module = _argListQualifiedContext.ModuleName.Component.CodeModule;
            module.DeleteLines(_subStmtQualifiedContext.Context.Start.Line,
                _subStmtQualifiedContext.Context.Stop.Line - _subStmtQualifiedContext.Context.Start.Line + 1);
            module.InsertLines(_subStmtQualifiedContext.Context.Start.Line, newfunctionWithReturn);
        }
    }
}