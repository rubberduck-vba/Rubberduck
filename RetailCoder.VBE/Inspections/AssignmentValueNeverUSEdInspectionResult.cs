using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    internal struct CursorIndex
    {
        private readonly int _line;
        private readonly int _column;

        internal int Line { get { return _line; } }
        internal int Column { get { return _column; } }

        internal CursorIndex(int line, int column)
        {
            _line = line;
            _column = column;
        }
    }

    public class AssignmentValueNeverUsedInspectionResult : CodeInspectionResultBase
    {
        private readonly IdentifierReference _target;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public AssignmentValueNeverUsedInspectionResult(IInspection inspection, IdentifierReference target)
            : base(inspection, string.Format(inspection.Description, target.IdentifierName), target.QualifiedModuleName, target.Context)
        {
            _target = target;
            _quickFixes = new[]
            {
                new RemoveUnusedAssignment(Context, GetSelection(), target.IdentifierName),
            };
        }

        private QualifiedSelection GetSelection()
        {
            var instructionBeginning = GetBeginningOfInstructionSelection(_target.QualifiedModuleName.Component.CodeModule, _target.Selection);
            var instructionEnd = GetEndOfInstructionSelection(_target.QualifiedModuleName.Component.CodeModule, _target.Selection);

            return new QualifiedSelection(QualifiedSelection.QualifiedName, new Selection(instructionBeginning.Line, instructionBeginning.Column, instructionEnd.Line, instructionEnd.Column));
        }

        private CursorIndex GetBeginningOfInstructionSelection(CodeModule module, Selection referenceSelection)
        {
            var currentLine = referenceSelection.StartLine;

            var codeLine = module.Lines[currentLine, 1].StripStringLiterals().Remove(referenceSelection.StartColumn);

            if (codeLine.LastIndexOf(':') == -1 &&
                !module.Lines[currentLine - 1, 1].EndsWith(" _"))
            {
                return new CursorIndex(currentLine, 1);
            }

            while (codeLine.LastIndexOf(':') == -1 && module.Lines[currentLine - 1, 1].EndsWith(" _"))
            {
                codeLine = module.Lines[--currentLine, 1].StripStringLiterals();
                if (!codeLine.EndsWith(" _"))
                {
                    return new CursorIndex(currentLine + 1, 1);
                }
            }

            var index = codeLine.LastIndexOf(':') == -1 ? 1 : codeLine.LastIndexOf(':');
            return new CursorIndex(currentLine, index);
        }

        private CursorIndex GetEndOfInstructionSelection(CodeModule module, Selection referenceSelection)
        {
            var currentLine = referenceSelection.EndLine;

            var codeLine = module.Lines[currentLine, 1].StripStringLiterals();

            // for the first line, make sure we are on the right instruction separator
            codeLine = codeLine.Remove(0, referenceSelection.EndColumn);
            for (var i = 0; i < referenceSelection.EndColumn; i++)
            {
                codeLine = codeLine.Insert(0, " ");
            }

            if (codeLine.IndexOf(':') == -1 &&
                !codeLine.EndsWith(" _"))
            {
                return new CursorIndex(currentLine, codeLine.Length);
            }

            while (codeLine.IndexOf(':') == -1 && codeLine.EndsWith(" _"))
            {
                codeLine = module.Lines[++currentLine, 1].StripStringLiterals();
                if (codeLine.IndexOf(':') != -1)
                {
                    var colonIndex = codeLine.IndexOf(':');
                    return new CursorIndex(currentLine, colonIndex);
                }

                if (!codeLine.EndsWith(" _"))
                {
                    return new CursorIndex(currentLine, codeLine.Length);
                }
            }

            var index = codeLine.IndexOf(':') == -1 ? codeLine.Length : codeLine.IndexOf(':');
            return new CursorIndex(currentLine, index);
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    /// <summary>
    /// Remove an instruction
    /// </summary>
    public class RemoveUnusedAssignment : CodeInspectionQuickFix
    {
        public RemoveUnusedAssignment(ParserRuleContext context, QualifiedSelection selection, string identifierName)
            : base(context, selection, string.Format(InspectionsUI.AssignmentValueNeverUsedInspectionQuickFix, identifierName))
        {
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            var lines = module.Lines[Selection.Selection.StartLine, Selection.Selection.LineCount].Split(new[] {Environment.NewLine}, StringSplitOptions.None);

            var newLines = lines.Length > 1
                ? lines[0].Remove(Selection.Selection.StartColumn) +
                  Environment.NewLine +
                  lines.Last().Remove(0, Selection.Selection.EndColumn)
                : lines[0].Remove(Selection.Selection.StartColumn,
                    Selection.Selection.EndColumn - Selection.Selection.StartColumn);

            module.DeleteLines(Selection.Selection.StartLine, Selection.Selection.LineCount);

            if (newLines.Trim() != string.Empty)
            {
                module.InsertLines(Selection.Selection.StartLine, newLines);
            }
        }
    }
}
