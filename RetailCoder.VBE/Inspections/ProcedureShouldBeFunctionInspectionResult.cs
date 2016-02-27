using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ProcedureShouldBeFunctionInspectionResult : InspectionResultBase
    {
       private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

       public ProcedureShouldBeFunctionInspectionResult(IInspection inspection, RubberduckParserState state, QualifiedContext<VBAParser.ArgListContext> argListQualifiedContext, QualifiedContext<VBAParser.SubStmtContext> subStmtQualifiedContext)
           : base(inspection,
                subStmtQualifiedContext.ModuleName,
                subStmtQualifiedContext.Context.ambiguousIdentifier())
        {
           _target = state.AllUserDeclarations.Single(declaration => 
               declaration.DeclarationType == DeclarationType.Procedure
               && declaration.Context == subStmtQualifiedContext.Context);

            _quickFixes = new[]
            {
                new ChangeProcedureToFunction(state, argListQualifiedContext, subStmtQualifiedContext, QualifiedSelection), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        private readonly Declaration _target;
        public override string Description
        {
            get { return string.Format(InspectionsUI.ProcedureCanBeWrittenAsFunctionInspectionResultFormat, _target.IdentifierName); }
        }
    }

    public class ChangeProcedureToFunction : CodeInspectionQuickFix
    {
        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }

        private readonly RubberduckParserState _state;
        private readonly QualifiedContext<VBAParser.ArgListContext> _argListQualifiedContext;
        private readonly QualifiedContext<VBAParser.SubStmtContext> _subStmtQualifiedContext;
        private readonly QualifiedContext<VBAParser.ArgContext> _argQualifiedContext;

        private int _lineOffset;

        public ChangeProcedureToFunction(RubberduckParserState state,
                                         QualifiedContext<VBAParser.ArgListContext> argListQualifiedContext,
                                         QualifiedContext<VBAParser.SubStmtContext> subStmtQualifiedContext,
                                         QualifiedSelection selection)
            : base(subStmtQualifiedContext.Context, selection, InspectionsUI.ProcedureShouldBeFunctionInspectionQuickFix)
        {
            _state = state;
            _argListQualifiedContext = argListQualifiedContext;
            _subStmtQualifiedContext = subStmtQualifiedContext;
            _argQualifiedContext = new QualifiedContext<VBAParser.ArgContext>(_argListQualifiedContext.ModuleName,
                _argListQualifiedContext.Context.arg()
                    .First(a => a.BYREF() != null || (a.BYREF() == null && a.BYVAL() == null)));
        }

        public override void Fix()
        {
            UpdateCalls();
            UpdateSignature();
        }

        private void UpdateSignature()
        {
            var argListText = _argListQualifiedContext.Context.GetText();
            var subStmtText = _subStmtQualifiedContext.Context.GetText();
            var argText = _argQualifiedContext.Context.GetText();

            var newArgText = argText.Contains("ByRef ") ? argText.Replace("ByRef ", "ByVal ") : "ByVal " + argText;

            var newFunctionWithoutReturn = subStmtText.Insert(subStmtText.IndexOf(argListText, StringComparison.Ordinal) + argListText.Length,
                                                              " " + _argQualifiedContext.Context.asTypeClause().GetText())
                                                      .Replace("Sub", "Function")
                                                      .Replace(argText, newArgText);

            var indexOfInstructionSeparators = new List<int>();
            var functionWithoutStringLiterals = newFunctionWithoutReturn.StripStringLiterals();
            for (var i = 0; i < functionWithoutStringLiterals.Length; i++)
            {
                if (functionWithoutStringLiterals[i] == ':')
                {
                    indexOfInstructionSeparators.Add(i);
                }
            }

            if (indexOfInstructionSeparators.Count > 1)
            {
                indexOfInstructionSeparators.Reverse();
                newFunctionWithoutReturn = indexOfInstructionSeparators.Aggregate(newFunctionWithoutReturn, (current, index) => current.Remove(index, 1).Insert(index, Environment.NewLine));
            }

            var newfunctionWithReturn = newFunctionWithoutReturn
                .Insert(newFunctionWithoutReturn.LastIndexOf(Environment.NewLine, StringComparison.Ordinal),
                        Environment.NewLine + "    " + _subStmtQualifiedContext.Context.ambiguousIdentifier().GetText() +
                        " = " + _argQualifiedContext.Context.ambiguousIdentifier().GetText());

            _lineOffset = newfunctionWithReturn.Split(new[] {Environment.NewLine}, StringSplitOptions.None).Length -
                         subStmtText.Split(new[] {Environment.NewLine}, StringSplitOptions.None).Length;

            var module = _argListQualifiedContext.ModuleName.Component.CodeModule;

            module.DeleteLines(_subStmtQualifiedContext.Context.Start.Line,
                _subStmtQualifiedContext.Context.Stop.Line - _subStmtQualifiedContext.Context.Start.Line + 1);
            module.InsertLines(_subStmtQualifiedContext.Context.Start.Line, newfunctionWithReturn);
        }

        private void UpdateCalls()
        {
            var procedureName = _subStmtQualifiedContext.Context.ambiguousIdentifier().GetText();

            var procedure =
                _state.AllDeclarations.SingleOrDefault(d =>
                        !d.IsBuiltIn &&
                        d.IdentifierName == procedureName &&
                        d.Context is VBAParser.SubStmtContext &&
                        d.ComponentName == _subStmtQualifiedContext.ModuleName.ComponentName &&
                        d.Project == _subStmtQualifiedContext.ModuleName.Project);

            if (procedure == null) { return; }

            foreach (var reference in procedure.References.OrderByDescending(o => o.Selection.StartLine).ThenByDescending(d => d.Selection.StartColumn))
            {
                var startLine = reference.Selection.StartLine;

                if (procedure.ComponentName == reference.QualifiedModuleName.ComponentName && procedure.Selection.EndLine < reference.Selection.StartLine)
                {
                    startLine += _lineOffset;
                }

                var module = reference.QualifiedModuleName.Component.CodeModule;

                var referenceParent = reference.Context.Parent as VBAParser.ICS_B_ProcedureCallContext;
                if (referenceParent == null) { continue; }
                
                var referenceText = reference.Context.Parent.GetText();
                var newCall = referenceParent.argsCall().argCall().ToList().ElementAt(_argListQualifiedContext.Context.arg().ToList().IndexOf(_argQualifiedContext.Context)).GetText() +
                              " = " + _subStmtQualifiedContext.Context.ambiguousIdentifier().GetText() +
                              "(" + referenceParent.argsCall().GetText() + ")";

                var oldLines = module.Lines[startLine, reference.Selection.LineCount];

                var newText = oldLines.Remove(reference.Selection.StartColumn - 1, referenceText.Length)
                    .Insert(reference.Selection.StartColumn - 1, newCall);

                module.DeleteLines(startLine, reference.Selection.LineCount);
                module.InsertLines(startLine, newText);
            }
        }
    }
}