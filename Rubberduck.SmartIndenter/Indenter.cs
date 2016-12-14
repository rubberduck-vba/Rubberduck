using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.SmartIndenter
{
    public class Indenter : IIndenter
    {
        private readonly IVBE _vbe;
        private readonly Func<IIndenterSettings> _settings;

        public Indenter(IVBE vbe, Func<IIndenterSettings> settings)
        {
            _vbe = vbe;
            _settings = settings;
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

            var procName = module.GetProcOfLine(selection.StartLine);
            var procKind = module.GetProcKindOfLine(selection.StartLine);

            if (string.IsNullOrEmpty(procName))
            {
                return;
            }

            var startLine = module.GetProcStartLine(procName, procKind);
            var endLine = startLine + module.GetProcCountLines(procName, procKind);

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
        
        private static Selection GetSelection(ICodePane codePane)
        {
            return codePane.Selection;
        }

        public void Indent(IVBComponent component)
        {
            var module = component.CodeModule;
            var lineCount = module.CountOfLines;
            if (lineCount == 0)
            {
                return;
            }

            var codeLines = module.GetLines(1, lineCount).Replace("\r", string.Empty).Split('\n');
            var indented = Indent(codeLines, component.Name);

            module.DeleteLines(1, lineCount);
            module.InsertLines(1, string.Join("\r\n", indented));
        }

        public void Indent(IVBComponent component, string procedureName, Selection selection)
        {
            var module = component.CodeModule;
            var lineCount = module.CountOfLines;
            if (lineCount == 0)
            {
                return;
            }

            var codeLines = module.GetLines(selection.StartLine, selection.LineCount).Replace("\r", string.Empty).Split('\n');

            var indented = Indent(codeLines, procedureName);

            var start = selection.StartLine;
            var lines = selection.LineCount;

            //Deletelines fails if the the last line of the procedure is the last line of the module.
            module.DeleteLines(start, start + lines < lineCount ? lines : lines - 1);
            module.InsertLines(start, string.Join("\r\n", indented));
        }

        private IEnumerable<LogicalCodeLine> BuildLogicalCodeLines(IEnumerable<string> lines)
        {
            var settings = _settings.Invoke();
            var logical = new List<LogicalCodeLine>();
            LogicalCodeLine current = null;
            AbsoluteCodeLine previous = null;

            foreach (var line in lines)
            {
                var absolute = new AbsoluteCodeLine(line, settings, previous);
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
                previous = absolute;
            }
            return logical;
        }

        public IEnumerable<string> Indent(IEnumerable<string> codeLines, string moduleName)
        {
            var logical = BuildLogicalCodeLines(codeLines).ToList();
            var indents = 0;
            var start = false;
            var enumStart = false;
            var inEnumType = false;

            foreach (var line in logical.Where(x => !x.IsEmpty))
            {
                inEnumType &= !line.IsEnumOrTypeEnd;
                if (inEnumType)
                {
                    line.AtEnumTypeStart = enumStart;
                    enumStart = line.IsCommentBlock;
                    line.IsEnumOrTypeMember = true;
                    line.IndentationLevel = line.EnumTypeIndents;
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
                enumStart = inEnumType;
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
