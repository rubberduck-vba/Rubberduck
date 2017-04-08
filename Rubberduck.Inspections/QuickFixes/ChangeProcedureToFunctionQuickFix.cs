using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ChangeProcedureToFunctionQuickFix : QuickFixBase
    {
        public override bool CanFixInModule
        {
            get { return false; }
        }

        public override bool CanFixInProject
        {
            get { return false; }
        }

        private readonly RubberduckParserState _state;
        private readonly QualifiedContext<VBAParser.ArgListContext> _argListQualifiedContext;
        private readonly QualifiedContext<VBAParser.SubStmtContext> _subStmtQualifiedContext;
        private readonly QualifiedContext<VBAParser.ArgContext> _argQualifiedContext;

        private int _lineOffset;

        public ChangeProcedureToFunctionQuickFix(RubberduckParserState state,
            QualifiedContext<VBAParser.ArgListContext> argListQualifiedContext,
            QualifiedContext<VBAParser.SubStmtContext> subStmtQualifiedContext,
            QualifiedSelection selection)
            : base(subStmtQualifiedContext.Context, selection, InspectionsUI.ProcedureShouldBeFunctionInspectionQuickFix
                )
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

            var asTypeClause = _argQualifiedContext.Context.asTypeClause() != null
                ? _argQualifiedContext.Context.asTypeClause().GetText()
                : "As Variant";

            var newFunctionWithoutReturn = subStmtText.Insert(
                subStmtText.IndexOf(argListText, StringComparison.Ordinal) + argListText.Length,
                " " + asTypeClause)
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
                newFunctionWithoutReturn = indexOfInstructionSeparators.Aggregate(newFunctionWithoutReturn,
                    (current, index) => current.Remove(index, 1).Insert(index, Environment.NewLine));
            }

            var newfunctionWithReturn = newFunctionWithoutReturn
                .Insert(newFunctionWithoutReturn.LastIndexOf(Environment.NewLine, StringComparison.Ordinal),
                    Environment.NewLine + "    " + _subStmtQualifiedContext.Context.subroutineName().GetText() +
                    " = " + _argQualifiedContext.Context.unrestrictedIdentifier().GetText());

            _lineOffset = newfunctionWithReturn.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Length -
                          subStmtText.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Length;

            var module = _argListQualifiedContext.ModuleName.Component.CodeModule;

            module.DeleteLines(_subStmtQualifiedContext.Context.Start.Line,
                _subStmtQualifiedContext.Context.Stop.Line - _subStmtQualifiedContext.Context.Start.Line + 1);
            module.InsertLines(_subStmtQualifiedContext.Context.Start.Line, newfunctionWithReturn);
        }

        private void UpdateCalls()
        {
            var procedureName = Identifier.GetName(_subStmtQualifiedContext.Context.subroutineName().identifier());

            var procedure =
                _state.AllUserDeclarations.SingleOrDefault(d =>
                    d.IdentifierName == procedureName &&
                    d.Context is VBAParser.SubStmtContext &&
                    d.QualifiedName.QualifiedModuleName.Equals(_subStmtQualifiedContext.ModuleName));

            if (procedure == null)
            {
                return;
            }

            foreach (
                var reference in
                    procedure.References.OrderByDescending(o => o.Selection.StartLine)
                        .ThenByDescending(d => d.Selection.StartColumn))
            {
                var startLine = reference.Selection.StartLine;

                if (procedure.ComponentName == reference.QualifiedModuleName.ComponentName &&
                    procedure.Selection.EndLine < reference.Selection.StartLine)
                {
                    startLine += _lineOffset;
                }

                var referenceParent = ParserRuleContextHelper.GetParent<VBAParser.CallStmtContext>(reference.Context);
                if (referenceParent == null)
                {
                    continue;
                }

                var module = reference.QualifiedModuleName.Component.CodeModule;
                {
                    var argList = CallStatement.GetArgumentList(referenceParent);
                    var paramNames = new List<string>();
                    var argsCall = string.Empty;
                    var argsCallOffset = 0;
                    if (argList != null)
                    {
                        argsCallOffset = argList.GetSelection().EndColumn - reference.Context.GetSelection().EndColumn;
                        argsCall = argList.GetText();
                        if (argList.argument() != null)
                        {
                            paramNames.AddRange(
                                argList.argument().Select(p =>
                                {
                                    if (p.positionalArgument() != null)
                                    {
                                        return p.positionalArgument().GetText();
                                    }
                                    if (p.namedArgument() != null)
                                    {
                                        return p.namedArgument().GetText();
                                    }
                                    return string.Empty;
                                }).ToList());
                        }
                    }
                    var referenceText = reference.Context.GetText();
                    var newCall =
                        paramNames.ToList()
                            .ElementAt(
                                _argListQualifiedContext.Context.arg().ToList().IndexOf(_argQualifiedContext.Context)) +
                        " = " + _subStmtQualifiedContext.Context.subroutineName().GetText() +
                        "(" + argsCall + ")";

                    var oldLines = module.GetLines(startLine, reference.Selection.LineCount);

                    var newText = oldLines.Remove(reference.Selection.StartColumn - 1,
                        referenceText.Length + argsCallOffset)
                        .Insert(reference.Selection.StartColumn - 1, newCall);

                    module.DeleteLines(startLine, reference.Selection.LineCount);
                    module.InsertLines(startLine, newText);
                }
            }
        }
    }
}