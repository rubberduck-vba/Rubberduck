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
            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane == null || pane.IsWrappingNullReference)
                {
                    return;
                }

                var initialSelection = GetSelection(pane).Collapse();

                using (var module = pane.CodeModule)
                {
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
                    using (var component = module.Parent)
                    {
                        Indent(component, selection, true);
                    }                   
                }
                ResetSelection(pane, initialSelection);
            }
        }

        /// <summary>
        /// Indents the code in the ActiveCodePane.
        /// </summary>
        public void IndentCurrentModule()
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane == null || pane.IsWrappingNullReference)
                {
                    return;
                }

                var initialSelection = GetSelection(pane).Collapse();

                using (var module = pane.CodeModule)
                {
                    using (var component = module.Parent)
                    {
                        Indent(component);
                    }                   
                }

                ResetSelection(pane, initialSelection);
            }
        }

        /// <summary>
        /// Indents every module in the active project.
        /// </summary>
        public void IndentCurrentProject()
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                var initialSelection = pane == null || pane.IsWrappingNullReference ? default : GetSelection(pane).Collapse();

                var project = _vbe.ActiveVBProject;
                if (project.Protection == ProjectProtection.Locked)
                {
                    return;
                }

                foreach (var component in project.VBComponents)
                {
                    Indent(component);
                }

                ResetSelection(pane, initialSelection);
            }
        }

        private void ResetSelection(ICodePane codePane, Selection initialSelection)
        {
            using (var window = _vbe.ActiveWindow)
            {
                if (initialSelection == default || codePane == null || window == null ||
                    window.IsWrappingNullReference || window.Type != WindowKind.CodeWindow ||
                    codePane.IsWrappingNullReference)
                {
                    return;
                }
            }

            using (var module = codePane.CodeModule)
            {
                // This will only "ballpark it" for now - it sets the absolute line in the module, not necessarily
                // the specific LoC. That will be a TODO when the parse tree is used to indent. For the time being,
                // maintaining that is ridiculously difficult vis-a-vis the payoff if the vertical spacing is 
                // changed.
                var lines = module.CountOfLines;
                codePane.Selection = lines < initialSelection.StartLine
                    ? new Selection(lines, initialSelection.StartColumn, lines, initialSelection.StartColumn)
                    : initialSelection;
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
            using (var module = component.CodeModule)
            {
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
        }

        /// <summary>
        /// Not fully implemented for selections (it does not track the current indentation level before the call). Use at your own
        /// risk on anything smaller than a procedure - the caller is responsible for determining the base indent and restoring it
        /// *after* the call.
        /// </summary>
        /// <param name="component">The VBComponent to indent</param>
        /// <param name="selection">The selection to indent</param>
        /// <param name="procedure">Whether the selection is a single procedure</param>
        private void Indent(IVBComponent component, Selection selection, bool procedure = false)
        {
            using (var module = component.CodeModule)
            {
                var lineCount = module.CountOfLines;
                if (lineCount == 0)
                {
                    return;
                }

                var codeLines = module.GetLines(selection.StartLine, selection.LineCount).Replace("\r", string.Empty)
                    .Split('\n');

                var indented = Indent(codeLines, false, procedure);

                var start = selection.StartLine;
                var lines = selection.LineCount;

                //Deletelines fails if the the last line of the procedure is the last line of the module.
                module.DeleteLines(start, start + lines < lineCount ? lines : lines - 1);
                module.InsertLines(start, string.Join("\r\n", indented));
            }
        }

        private IEnumerable<LogicalCodeLine> BuildLogicalCodeLines(IEnumerable<string> lines, out IIndenterSettings settings)
        {
            settings = _settings.Invoke();
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
        /// Indents the code contained in the passed string. NOTE: This overload should only be used on procedures or modules.
        /// </summary>
        /// <param name="code">The code block to indent</param>
        /// <returns>Indented code lines</returns>
        public IEnumerable<string> Indent(string code)
        {
            return Indent(code.Replace("\r", string.Empty).Split('\n'), false);
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
            return Indent(codeLines, forceTrailingNewLines, false);
        }

        private IEnumerable<string> Indent(IEnumerable<string> codeLines, bool forceTrailingNewLines, bool procedure)
        {
            var logical = BuildLogicalCodeLines(codeLines, out var settings).ToList();
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
                start = line.IsProcedureStart || 
                        line.AtProcedureStart && line.IsDeclaration ||
                        line.AtProcedureStart && line.IsCommentBlock ||
                        settings.IgnoreEmptyLinesInFirstBlocks && line.AtProcedureStart && line.IsEmpty;
                inEnumType = line.IsEnumOrTypeStart;
                enumStart = inEnumType;
            }

            return GenerateCodeLineStrings(logical, forceTrailingNewLines, procedure);
        }

        private IEnumerable<string> GenerateCodeLineStrings(IEnumerable<LogicalCodeLine> logical, bool forceTrailingNewLines, bool procedure = false)
        {
            var output = new List<string>();
            var settings = _settings.Invoke();

            List<LogicalCodeLine> indent;
            if (!procedure && settings.VerticallySpaceProcedures)
            {               
                indent = new List<LogicalCodeLine>();
                var lines = logical.ToArray();
                var header = true;
                var inEnumType = false;
                for (var i = 0; i < lines.Length; i++)
                {
                    indent.Add(lines[i]);

                    if (header && lines[i].IsEnumOrTypeStart)
                    {
                        inEnumType = true;
                    }
                    if (header && lines[i].IsEnumOrTypeEnd)
                    {
                        inEnumType = false;
                    }

                    if (header && !inEnumType && lines[i].IsProcedureStart)
                    {
                        header = false;
                        SpaceHeader(indent, settings);
                        continue;
                    }
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
                    else if (forceTrailingNewLines && i == lines.Length)
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

        private static void SpaceHeader(IList<LogicalCodeLine> header, IIndenterSettings settings)
        {
            var commentSkipped = false;
            var commentLines = 0;
            for (var i = header.Count - 2; i >= 0; i--)
            {
                if (!commentSkipped && header[i].IsCommentBlock)
                {
                    commentLines++;
                    continue;
                }

                commentSkipped = true;
                if (header[i].IsEmpty)
                {
                    header.RemoveAt(i);
                }
                else
                {
                    header.Insert(header.Count - 1 - commentLines,
                        new LogicalCodeLine(
                            Enumerable.Repeat(new AbsoluteCodeLine(string.Empty, settings),
                                settings.LinesBetweenProcedures), settings));
                    return;
                }
            }
        }
    }
}
