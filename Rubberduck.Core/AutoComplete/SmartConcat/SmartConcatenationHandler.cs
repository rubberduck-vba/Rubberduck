using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.AutoComplete.SmartConcat
{
    /// <summary>
    /// Adds a line continuation when {ENTER} is pressed when inside a string literal.
    /// </summary>
    public class SmartConcatenationHandler : AutoCompleteHandlerBase
    {
        public SmartConcatenationHandler(ICodePaneHandler pane) 
            : base(pane)
        {
        }

        public override bool Handle(AutoCompleteEventArgs e, AutoCompleteSettings settings, out CodeString result)
        {
            result = null;
            if (e.Character != '\r' || (!settings?.SmartConcat.IsEnabled ?? true))
            {
                return false;
            }

            var currentContent = CodePaneHandler.GetCurrentLogicalLine(e.Module);
            if ((!currentContent?.IsInsideStringLiteral ?? true) 
                || currentContent.Lines.Length >= settings.SmartConcat.ConcatMaxLines)
            {
                // selection spans more than a single logical line, or spans too many lines to be legal;
                // too many line continuations throws COMException if we attempt to modify.
                return false;
            }

            var lastIndexLeftOfCaret = currentContent.CaretLine.Length > 2 ? currentContent.CaretLine.Substring(0, currentContent.CaretPosition.StartColumn).LastIndexOf('"') : 0;
            if (lastIndexLeftOfCaret > 0)
            {
                var indent = currentContent.CaretLine.NthIndexOf('"', 1);
                var whitespace = new string(' ', indent);

                // todo: handle shift modifier?
                var concatVbNewLine = settings.SmartConcat.ConcatVbNewLineModifier.HasFlag(ModifierKeySetting.CtrlKey) && e.IsControlKeyDown;

                var autoCode = $"\" {(concatVbNewLine ? "& vbNewLine " : string.Empty)}& _\r\n{whitespace}\"";
                var left = currentContent.CaretLine.Substring(0, currentContent.CaretPosition.StartColumn);
                var right = currentContent.CaretLine.Substring(currentContent.CaretPosition.StartColumn);

                var caretLine = $"{left}{autoCode}{right}";
                var lines = currentContent.Lines;
                lines[currentContent.CaretPosition.StartLine] = caretLine;
                var code = string.Join("\r\n", lines);

                var newContent = new CodeString(code, currentContent.CaretPosition, currentContent.SnippetPosition);
                var newPosition = new Selection(newContent.CaretPosition.StartLine + 1, indent + 1);

                e.Handled = true;
                result = new CodeString(newContent.Code, newPosition, 
                    new Selection(newContent.SnippetPosition.StartLine, 1, newContent.SnippetPosition.EndLine, 1));

                CodePaneHandler.SubstituteCode(e.Module, result);
                var finalSelection = new Selection(result.SnippetPosition.StartLine, 1).Offset(result.CaretPosition);
                CodePaneHandler.SetSelection(e.Module, finalSelection);
                return true;
            }

            return false;
        }
    }
}