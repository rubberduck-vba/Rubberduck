using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.SmartIndenter
{
    public class Indenter : IIndenter
    {
        private readonly VBE _vbe;
        private readonly Func<IIndenterSettings> _settings;

        public Indenter(VBE vbe, Func<IIndenterSettings> settings)
        {
            _vbe = vbe;
            _settings = settings;
        }

        public event EventHandler<IndenterProgressEventArgs> ReportProgress;

        //TODO: Unimplemented.
        // ReSharper disable once UnusedMember.Local
        private void OnReportProgress(string moduleName, int progress, int max)
        {
            var handler = ReportProgress;
            if (handler == null) return;
            var args = new IndenterProgressEventArgs(moduleName, progress, max);
            handler.Invoke(this, args);
        }

        public void IndentCurrentProcedure()
        {
            var pane = _vbe.ActiveCodePane;

            if (pane == null)
            {
                return;
            }
            var module = pane.CodeModule;
            var selection = GetSelection(pane);

            vbext_ProcKind procKind;
            // ReSharper disable once UseIndexedProperty
            var procName = module.get_ProcOfLine(selection.StartLine, out procKind);

            if (string.IsNullOrEmpty(procName))
            {
                return;
            }

            var startLine = module.ProcStartLine[procName, procKind];
            var endLine = startLine + module.ProcCountLines[procName, procKind];

            selection = new Selection(startLine, 1, endLine, 1);
            Indent(module.Parent, procName, selection);
        }

        public void IndentCurrentModule()
        {
            var pane = _vbe.ActiveCodePane;
            if (pane == null)
            {
                return;
            }
            Indent(pane.CodeModule.Parent);
        }

        private static Selection GetSelection(CodePane codePane)
        {
            int startLine;
            int startColumn;
            int endLine;
            int endColumn;
            codePane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
            return new Selection(startLine, startColumn, endLine, endColumn);
        }

        public void Indent(VBComponent component, bool reportProgress = true, int linesAlreadyRebuilt = 0)
        {
            var module = component.CodeModule;
            var lineCount = module.CountOfLines;
            if (lineCount == 0)
            {
                return;
            }

            var codeLines = module.Lines[1, lineCount].Replace("\r", string.Empty).Split('\n');
            var indented = Indent(codeLines, component.Name, reportProgress, linesAlreadyRebuilt).ToArray();

            for (var i = 0; i < lineCount; i++)
            {
                if (module.Lines[i + 1, 1] != indented[i])
                {
                    component.CodeModule.ReplaceLine(i + 1, indented[i]);
                }
            }
        }

        public void Indent(VBComponent component, string procedureName, Selection selection, bool reportProgress = true, int linesAlreadyRebuilt = 0)
        {
            var module = component.CodeModule;
            var lineCount = module.CountOfLines;
            if (lineCount == 0)
            {
                return;
            }

            var codeLines = module.Lines[selection.StartLine, selection.LineCount].Replace("\r", string.Empty).Split('\n');

            var indented = Indent(codeLines, procedureName, reportProgress, linesAlreadyRebuilt).ToArray();

            for (var i = 0; i < selection.EndLine - selection.StartLine; i++)
            {
                if (module.Lines[selection.StartLine + i, 1] != indented[i])
                {
                    component.CodeModule.ReplaceLine(selection.StartLine + i, indented[i]);
                }
            }
        }

        private IEnumerable<LogicalCodeLine> BuildLogicalCodeLines(IEnumerable<string> lines)
        {
            var settings = _settings.Invoke();
            var logical = new List<LogicalCodeLine>();
            LogicalCodeLine current = null;

            foreach (var line in lines)
            {
                var absolute = new AbsoluteCodeLine(line, settings);
                if (current == null)
                {
                    current = new LogicalCodeLine(absolute, settings);
                    logical.Add(current);
                }
                else
                {
                    current.AddContinuationLine(absolute);
                }

                if (!absolute.HasContinuation)
                {
                    current = null;
                }
            }
            return logical;
        }

        public IEnumerable<string> Indent(IEnumerable<string> codeLines, string moduleName, bool reportProgress = true, int linesAlreadyRebuilt = 0)
        {
            var logical = BuildLogicalCodeLines(codeLines).ToList();
            var indents = 0;
            var start = false;
            var inEnumType = false;

            foreach (var line in logical.Where(x => !x.IsEmpty))
            {
                inEnumType &= !line.IsEnumOrTypeEnd;
                if (inEnumType)
                {
                    line.IsEnumOrTypeMember = true;
                    line.IndentationLevel = 1;
                    continue;
                }
                if (line.IsProcedureStart || line.IsEnumOrTypeStart)
                {
                    indents = 0;
                }
                line.AtProcedureStart = start;
                line.IndentationLevel = indents - line.Outdents;
                indents += line.NextLineIndents;
                start = line.IsProcedureStart || (line.AtProcedureStart && line.IsDeclaration) || (line.AtProcedureStart && line.IsCommentBlock);
                inEnumType = line.IsEnumOrTypeStart;
            }

            var output = new List<string>();
            foreach (var line in logical)
            {
                output.AddRange(line.Indented().Split(new[] { Environment.NewLine }, StringSplitOptions.None));
            }
            return output;
        }
    }
}
