using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.AutoComplete.Service
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

        public override CodeString Handle(AutoCompleteEventArgs e, AutoCompleteSettings settings)
        {
            if (e.Character != '\r' || (!settings?.EnableSmartConcat ?? true))
            {
                return null;
            }

            var currentContent = CodePaneHandler.GetCurrentLogicalLine(e.Module);
            var shouldHandle = currentContent.IsInsideStringLiteral;
            if (!shouldHandle)
            {
                return null;
            }

            var lastIndexLeftOfCaret = currentContent.Code.Length > 2 ? currentContent.Code.Substring(0, currentContent.CaretCharIndex).LastIndexOf('"') : 0;
            if (lastIndexLeftOfCaret > 0)
            {
                var indent = currentContent.Code.NthIndexOf('"', 1);
                var whitespace = new string(' ', indent);

                var autoCode = $"\" {(e.IsControlKeyDown ? " & vbNewLine " : string.Empty)} & _\r\n{whitespace}\"";
                var code = $"{currentContent.Code.Substring(0, currentContent.CaretCharIndex)}{autoCode}{currentContent.Code.Substring(currentContent.CaretCharIndex + 1)}";

                var newContent = new CodeString(code, currentContent.CaretPosition, currentContent.SnippetPosition);
                var newPosition = new Selection(newContent.CaretPosition.StartLine + 1, indent + 1);

                e.Handled = true;
                var result = new CodeString(newContent.Code, newPosition, 
                    new Selection(newContent.SnippetPosition.StartLine, 1, newContent.SnippetPosition.EndLine, 1));

                CodePaneHandler.SubstituteCode(e.Module, result);
                CodePaneHandler.SetSelection(e.Module, result.SnippetPosition.Offset(result.CaretPosition));
                return result;
            }

            return null;
        }
    }
}