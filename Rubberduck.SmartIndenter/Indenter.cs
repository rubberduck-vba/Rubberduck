using System;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.SmartIndenter
{
    public sealed class Indenter : SimpleIndenter, IIndenter
    {
        private readonly IVBE _vbe;
        private readonly Func<IIndenterSettings> _settings;

        public Indenter(IVBE vbe, Func<IIndenterSettings> settings)
        {
            _vbe = vbe;
            _settings = settings;
        }

        protected override Func<IIndenterSettings> Settings => _settings;

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
                var indented = Indent(codeLines, true, _settings.Invoke());

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

                var indented = Indent(codeLines, false, procedure, _settings.Invoke());

                var start = selection.StartLine;
                var lines = selection.LineCount;

                //Deletelines fails if the the last line of the procedure is the last line of the module.
                module.DeleteLines(start, start + lines < lineCount ? lines : lines - 1);
                module.InsertLines(start, string.Join("\r\n", indented));
            }
        }
    }
}
