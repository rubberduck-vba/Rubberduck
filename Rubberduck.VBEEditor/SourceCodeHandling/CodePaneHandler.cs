using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Diagnostics;

namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public class CodePaneHandler : ICodePaneHandler
    {
        private readonly IProjectsProvider _projectsProvider;

        public CodePaneHandler(IProjectsProvider projectsProvider)
        {
            _projectsProvider = projectsProvider;
        }

        public string SourceCode(QualifiedModuleName module)
        {
            var component = _projectsProvider.Component(module);
            if (component == null)
            {
                return string.Empty;
            }

            using (var codeModule = component.CodeModule)
            {
                return codeModule.Content() ?? string.Empty;
            }
        }

        public void SetSelection(ICodeModule module, Selection selection)
        {
            using (var pane = module.CodePane)
            {
                pane.Selection = selection;
            }
        }

        public void SubstituteCode(ICodeModule module, CodeString newCode)
        {
            try
            {
                module.DeleteLines(newCode.SnippetPosition);
                module.InsertLines(newCode.SnippetPosition.StartLine, newCode.Code);
            }
            catch
            {
                Debug.Assert(false, "too many line continuations. we shouldn't even be here.");
            }
        }

        public void SubstituteCode(QualifiedModuleName module, CodeString newCode)
        {
            var component = _projectsProvider.Component(module);
            if (component == null)
            {
                return;
            }

            using (var codeModule = component.CodeModule)
            {
                SubstituteCode(codeModule, newCode);
            }
        }

        public void SubstituteCode(QualifiedModuleName module, string newCode)
        {
            var component = _projectsProvider.Component(module);
            if (component == null)
            {
                return;
            }

            using (var codeModule = component.CodeModule)
            {
                codeModule.Clear();
                codeModule.InsertLines(1, newCode);
            }
        }

        public CodeString Prettify(ICodeModule module, CodeString original)
        {
            var originalCode = original.Code.Replace("\r", string.Empty).Split('\n');
            var originalPosition = original.CaretPosition.StartColumn;
            var isAtLastCharacter = originalPosition == original.CaretLine.Length;

            var originalNonWhitespaceCharacters = 0;
            var isAllWhitespace = !isAtLastCharacter;
            if (!isAtLastCharacter)
            {
                for (var i = 0; i <= Math.Min(originalPosition - 1, original.CaretLine.Length - 1); i++)
                {
                    if (originalCode[original.CaretPosition.StartLine][i] != ' ')
                    {
                        originalNonWhitespaceCharacters++;
                        isAllWhitespace = false;
                    }
                }
            }

            var indent = original.CaretLine.TakeWhile(c => c == ' ').Count();

            module.DeleteLines(original.SnippetPosition.StartLine, original.SnippetPosition.LineCount);
            module.InsertLines(original.SnippetPosition.StartLine, string.Join("\r\n", originalCode));

            var prettifiedCode = module.GetLines(original.SnippetPosition)
                                           .Replace("\r", string.Empty)
                                           .Split('\n');

            var prettifiedNonWhitespaceCharacters = 0;
            var prettifiedCaretCharIndex = 0;
            if (!isAtLastCharacter)
            {
                for (var i = 0; i < prettifiedCode[original.CaretPosition.StartLine].Length; i++)
                {
                    if (prettifiedCode[original.CaretPosition.StartLine][i] != ' ')
                    {
                        prettifiedNonWhitespaceCharacters++;
                        if (prettifiedNonWhitespaceCharacters == originalNonWhitespaceCharacters
                            || i == prettifiedCode[original.CaretPosition.StartLine].Length - 1)
                        {
                            prettifiedCaretCharIndex = i;
                            break;
                        }
                    }
                }
            }
            else
            {
                prettifiedCaretCharIndex = prettifiedCode[original.CaretPosition.StartLine].Length;
            }

            var prettifiedPosition = new Selection(
                    original.SnippetPosition.ToZeroBased().StartLine + original.CaretPosition.StartLine,
                    prettifiedCode[original.CaretPosition.StartLine].Trim().Length == 0 || (isAllWhitespace && !string.IsNullOrEmpty(original.CaretLine.Substring(original.CaretPosition.StartColumn).Trim()))
                        ? Math.Min(indent, original.CaretPosition.StartColumn)
                        : Math.Min(prettifiedCode[original.CaretPosition.StartLine].Length, prettifiedCaretCharIndex + 1))
                .ToOneBased();

            SetSelection(module, prettifiedPosition);

            return GetPrettifiedCodeString(original, prettifiedPosition, prettifiedCode);
        }

        public CodeString Prettify(QualifiedModuleName module, CodeString original)
        {
            var component = _projectsProvider.Component(module);
            if (component == null)
            {
                return original;
            }

            using (var codeModule = component.CodeModule)
            {
                return Prettify(codeModule, original);
            }
        }

        private static CodeString GetPrettifiedCodeString(CodeString original, Selection prettifiedPosition, string[] prettifiedCode)
        {
            var caretPosition = new Selection(original.CaretPosition.StartLine,
                prettifiedPosition.StartColumn - 1); // caretPosition is zero-based

            var snippetPosition = new Selection(original.SnippetPosition.StartLine,
                original.SnippetPosition.StartColumn, original.SnippetPosition.EndLine,
                prettifiedCode[prettifiedCode.Length - 1].Length);

            var result = new CodeString(string.Join("\r\n", prettifiedCode), caretPosition, snippetPosition);
            return result;
        }

        private const string LineContinuation = " _";

        public CodeString GetCurrentLogicalLine(ICodeModule module)
        {
            Selection pSelection;
            using (var pane = module.CodePane)
            {
                pSelection = pane.Selection;
            }

            var selectedContent = module.GetLines(pSelection.StartLine, pSelection.LineCount);
            var selectedLines = selectedContent.Replace("\r", string.Empty).Split('\n');
            var currentLine = selectedLines[0];

            var caretStartLine = (pSelection.StartLine, currentLine);
            var lines = new List<(int pLine, string Content)> {caretStartLine};

            // selection line may not be the only physical line in the complete logical line; accounts for line continuations.
            InsertPhysicalLinesAboveSelectionStart(lines, module, pSelection.StartLine);
            AppendPhysicalLinesBelowSelectionStart(lines, module, pSelection.StartLine, currentLine);

            var logicalLine = string.Join("\r\n", lines.Select(e => e.Content));

            var zCaretLine = lines.IndexOf(caretStartLine);
            var zCaretColumn = pSelection.StartColumn - 1;
            var caretPosition = new Selection(
                zCaretLine, zCaretColumn, zCaretLine + pSelection.LineCount - 1, pSelection.EndColumn - 1);

            var pStartLine = lines[0].pLine;
            var pEndLine = lines[lines.Count - 1].pLine;
            var snippetPosition = new Selection(pStartLine, 1, pEndLine, 1);

            if (pStartLine > pSelection.StartLine || pEndLine > pSelection.EndLine)
            {
                // selection spans more than a single logical line
                return null;
            }

            var result = new CodeString(logicalLine, caretPosition, snippetPosition);
            return result;
        }

        private void AppendPhysicalLinesBelowSelectionStart(ICollection<(int Line, string Content)> lines, ICodeModule module, int currentLineIndex, string currentLine)
        {
            // assumes caret line is already in the list.
            while (currentLineIndex <= module.CountOfLines && currentLine.EndsWith(LineContinuation))
            {
                currentLineIndex++;
                if (currentLineIndex <= module.CountOfLines)
                {
                    currentLine = module.GetLines(currentLineIndex, 1);
                    lines.Add((currentLineIndex, currentLine));
                }
                else
                {
                    break;
                }
            }
        }

        private void InsertPhysicalLinesAboveSelectionStart(IList<(int Line, string Content)> lines, ICodeModule module, int currentLineIndex)
        {
            // assumes caret line is already in the list.
            while (currentLineIndex >= 1)
            {
                currentLineIndex--;
                if (currentLineIndex >= 1)
                {
                    var currentLine = module.GetLines(currentLineIndex, 1);
                    if (currentLine.EndsWith(LineContinuation))
                    {
                        lines.Insert(0, (currentLineIndex, currentLine));
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }
    }
}
