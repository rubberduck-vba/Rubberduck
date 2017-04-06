using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ChangeProcedureToFunctionQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type> { typeof(ProcedureCanBeWrittenAsFunctionInspection) };
        public static IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public static void AddSupportedInspectionType(Type inspectionType)
        {
            if (!inspectionType.GetInterfaces().Contains(typeof(IInspection)))
            {
                throw new ArgumentException("Type must implement IInspection", nameof(inspectionType));
            }

            _supportedInspections.Add(inspectionType);
        }

        private readonly RubberduckParserState _state;

        public ChangeProcedureToFunctionQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public void Fix(IInspectionResult result)
        {
            var subStmt = (VBAParser.SubStmtContext) result.Target.Context;
            var parameterList = ((VBAParser.SubStmtContext) result.Target.Context).argList();// new QualifiedContext<VBAParser.SubStmtContext>(result.Target.QualifiedName, (VBAParser.SubStmtContext)result.Target.Context);
            var arg = parameterList.arg().First(a => a.BYREF() != null || (a.BYREF() == null && a.BYVAL() == null));

            UpdateCalls(subStmt, parameterList, arg, result.Target.QualifiedName.QualifiedModuleName);
            UpdateSignature(subStmt, parameterList, arg, result.Target.QualifiedName.QualifiedModuleName);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.ProcedureShouldBeFunctionInspectionQuickFix;
        }

        private void UpdateSignature(VBAParser.SubStmtContext subStmt, VBAParser.ArgListContext parameterList, VBAParser.ArgContext arg, QualifiedModuleName qualifiedModuleName)
        {
            var argListText = parameterList.GetText();
            var subStmtText = subStmt.GetText();
            var argText = arg.GetText();

            var newArgText = argText.Contains("ByRef ") ? argText.Replace("ByRef ", "ByVal ") : "ByVal " + argText;

            var asTypeClause = arg.asTypeClause() != null
                ? arg.asTypeClause().GetText()
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
                    Environment.NewLine + "    " + subStmt.subroutineName().GetText() +
                    " = " + arg.unrestrictedIdentifier().GetText());

            var module = qualifiedModuleName.Component.CodeModule;

            module.DeleteLines(subStmt.Start.Line,
                subStmt.Stop.Line - subStmt.Start.Line + 1);
            module.InsertLines(subStmt.Start.Line, newfunctionWithReturn);
        }

        private void UpdateCalls(VBAParser.SubStmtContext subStmt, VBAParser.ArgListContext parameterList, VBAParser.ArgContext arg, QualifiedModuleName qualifiedModuleName)
        {
            var procedureName = Identifier.GetName(subStmt.subroutineName().identifier());

            var procedure =
                _state.AllUserDeclarations.SingleOrDefault(d =>
                    d.IdentifierName == procedureName &&
                    d.Context is VBAParser.SubStmtContext &&
                    d.QualifiedName.QualifiedModuleName.Equals(qualifiedModuleName));

            if (procedure == null)
            {
                return;
            }

            foreach (var reference in
                    procedure.References.OrderByDescending(o => o.Selection.StartLine)
                        .ThenByDescending(d => d.Selection.StartColumn))
            {
                var startLine = reference.Selection.StartLine;

                var referenceParent = ParserRuleContextHelper.GetParent<VBAParser.CallStmtContext>(reference.Context);
                if (referenceParent == null)
                {
                    continue;
                }

                var module = reference.QualifiedModuleName.Component.CodeModule;
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
                            parameterList.arg().ToList().IndexOf(arg)) +
                    " = " + subStmt.subroutineName().GetText() +
                    "(" + argsCall + ")";

                var oldLines = module.GetLines(startLine, reference.Selection.LineCount);

                var newText = oldLines.Remove(reference.Selection.StartColumn - 1,
                        referenceText.Length + argsCallOffset)
                    .Insert(reference.Selection.StartColumn - 1, newCall);

                module.DeleteLines(startLine, reference.Selection.LineCount);
                module.InsertLines(startLine, newText);
            }
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => false;
        public bool CanFixInProject => false;
    }
}