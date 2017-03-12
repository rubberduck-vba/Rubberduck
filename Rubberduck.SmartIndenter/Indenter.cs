using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
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

        /// <summary>
        /// Indents the procedure selected in the ActiveCodePane. If more than one is selected, the first is indented.
        /// </summary>
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
            Indent(module.Parent, selection);
        }

        /// <summary>
        /// Indents the code in the ActiveCodePane.
        /// </summary>
        public void IndentCurrentModule()
        {
            var pane = _vbe.ActiveCodePane;
            if (pane == null)
            {
                return;
            }
            Indent(pane.CodeModule.Parent);
        }

        /// <summary>
        /// Indents every module in the active project.
        /// </summary>
        public void IndentCurrentProject()
        {
            var project = _vbe.ActiveVBProject;
            if (project.Protection == ProjectProtection.Locked)
            {
                return;
            }
            foreach (var component in project.VBComponents)
            {
                Indent(component);
            }
        }

        private static Selection GetSelection(ICodePane codePane)
        {
            return codePane.Selection;
        }

        /// <summary>
        /// Indents the code in the VBComponent's CodeModule.
        /// </summary>
        /// <param name="component">The VBComponent to indent</param>
        public void Indent(IVBComponent component)
        {
            var module = component.CodeModule;
            var lineCount = module.CountOfLines;
            if (lineCount == 0)
            {
                return;
            }

            var codeLines = module.GetLines(1, lineCount).Replace("\r", string.Empty).Split('\n');
            var indented = Indent(codeLines, true);

            module.DeleteLines(1, lineCount);
            module.InsertLines(1, string.Join("\r\n", indented));
        }

        /// <summary>
        /// DO NOT USE - Not fully implemented. Use the Indent(IVBComponent component) instead and ping @Comintern if you need this functionality...
        /// </summary>
        /// <param name="component">The VBComponent to indent</param>
        /// <param name="selection">The selection to indent</param>
        public void Indent(IVBComponent component, Selection selection)
        {
            var module = component.CodeModule;
            var lineCount = module.CountOfLines;
            if (lineCount == 0)
            {
                return;
            }

            var codeLines = module.GetLines(selection.StartLine, selection.LineCount).Replace("\r", string.Empty).Split('\n');

            var indented = Indent(codeLines);

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

        /// <summary>
        /// Indents a range of code lines. NOTE: If inserting procedures, use the forceTrailingNewLines overload to preserve vertical spacing in the module.
        /// Do not call directly on selections. Use Indent(IVBComponent, Selection) instead.
        /// </summary>
        /// <param name="codeLines">Code lines to indent</param>
        /// <returns>Indented code lines</returns>
        public IEnumerable<string> Indent(IEnumerable<string> codeLines)
        {
            return Indent(codeLines, false);
        }

        /// <summary>
        /// Indents a range of code lines. Do not call directly on selections. Use Indent(IVBComponent, Selection) instead.
        /// </summary>
        /// <param name="codeLines">Code lines to indent</param>
        /// <param name="forceTrailingNewLines">If true adds a number of blank lines after the last procedure based on VerticallySpaceProcedures settings</param>
        /// <returns>Indented code lines</returns>
        public IEnumerable<string> Indent(IEnumerable<string> codeLines, bool forceTrailingNewLines)
        {
            var logical = BuildLogicalCodeLines(codeLines).ToList();
            var indents = 0;
            var start = false;
            var enumStart = false;
            var inEnumType = false;
            var inProcedure = false;

            foreach (var line in logical)
            {
                inEnumType &= !line.IsEnumOrTypeEnd;
                if (inEnumType)
                {
                    line.AtEnumTypeStart = enumStart;
                    enumStart = line.IsCommentBlock;
                    line.IsEnumOrTypeMember = true;
                    line.InsideProcedureTypeOrEnum = true;
                    line.IndentationLevel = line.EnumTypeIndents;                    
                    continue;
                }

                if (line.IsProcedureStart)
                {
                    inProcedure = true;                    
                }                               
                line.InsideProcedureTypeOrEnum = inProcedure || enumStart;
                inProcedure = inProcedure && !line.IsProcudureEnd && !line.IsEnumOrTypeEnd;
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

            return GenerateCodeLineStrings(logical, forceTrailingNewLines);
        }

        private IEnumerable<string> GenerateCodeLineStrings(IEnumerable<LogicalCodeLine> logical, bool forceTrailingNewLines)
        {
            var output = new List<string>();
            var settings = _settings.Invoke();

            List<LogicalCodeLine> indent;
            if (settings.VerticallySpaceProcedures)
            {
                indent = new List<LogicalCodeLine>();
                var lines = logical.ToArray();
                for (var i = 0; i < lines.Length; i++)
                {
                    indent.Add(lines[i]);
                    if (!lines[i].IsEnumOrTypeEnd && !lines[i].IsProcudureEnd)
                    {
                        continue;
                    }
                    while (++i < lines.Length && lines[i].IsEmpty) { }
                    if (i != lines.Length)
                    {
                        if (settings.LinesBetweenProcedures > 0)
                        {
                            indent.Add(new LogicalCodeLine(Enumerable.Repeat(new AbsoluteCodeLine(string.Empty, settings), settings.LinesBetweenProcedures), settings));
                        }
                        indent.Add(lines[i]);
                    }
                    else if (i == lines.Length && forceTrailingNewLines)
                    {
                        indent.Add(new LogicalCodeLine(Enumerable.Repeat(new AbsoluteCodeLine(string.Empty, settings), Math.Max(settings.LinesBetweenProcedures, 1)), settings));
                    }
                }
            }
            else
            {
                indent = logical.ToList();
            }

            foreach (var line in indent)
            {
                output.AddRange(line.Indented().Split(new[] { Environment.NewLine }, StringSplitOptions.None));
            }
            return output;
        }
    }
}
