using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public class CodePaneSourceCodeHandler : ICodePaneHandler
    {
        private readonly IProjectsProvider _projectsProvider;

        public CodePaneSourceCodeHandler(IProjectsProvider projectsProvider)
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

        public void SetSelection(QualifiedModuleName module, Selection selection)
        {
            var component = _projectsProvider.Component(module);
            if (component == null)
            {
                return;
            }

            using (var codeModule = component.CodeModule)
            {
                SetSelection(codeModule, selection);
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
            module.DeleteLines(newCode.SnippetPosition);
            module.InsertLines(newCode.SnippetPosition.StartLine, newCode.Code);
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
            var originalNonWhitespaceCharacters = 0;
            var isAllWhitespace = true;
            for (var i = 0; i <= Math.Min(originalPosition - 1, originalCode[original.CaretPosition.StartLine].Length - 1); i++)
            {
                if (originalCode[original.CaretPosition.StartLine][i] != ' ')
                {
                    originalNonWhitespaceCharacters++;
                    isAllWhitespace = false;
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

        public CodeString GetCurrentLogicalLine(ICodeModule module)
        {
            const string lineContinuation = " _";

            Selection pSelection;
            using (var pane = module.CodePane)
            {
                pSelection = pane.Selection;
            }

            var currentLineIndex = pSelection.StartLine;
            var currentLine = module.GetLines(currentLineIndex, 1);

            var caretLine = (currentLineIndex, currentLine);
            var lines = new List<(int Line, string Content)> {caretLine};

            while (currentLineIndex >= 1)
            {
                currentLineIndex--;
                if (currentLineIndex >= 1)
                {
                    currentLine = module.GetLines(currentLineIndex, 1);
                    if (currentLine.EndsWith(lineContinuation))
                    {
                        lines.Insert(0, (currentLineIndex, currentLine));
                    }
                    else
                    {
                        break;
                    }
                }
            }

            currentLineIndex = pSelection.StartLine;
            currentLine = caretLine.currentLine;
            while (currentLineIndex <= module.CountOfLines && currentLine.EndsWith(lineContinuation))
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

            var logicalLine = string.Join("\r\n", lines.Select(e => e.Content));
            var zCaretLine = lines.IndexOf(caretLine);
            var zCaretColumn = pSelection.StartColumn - 1;

            var startLine = lines[0].Line;
            var endLine = lines[lines.Count - 1].Line;

            var result = new CodeString(
                logicalLine,
                new Selection(zCaretLine, zCaretColumn),
                new Selection(startLine, 1, endLine, 1));

            return result;

        }

        public CodeString GetCurrentLogicalLine(QualifiedModuleName module)
        {
            var component = _projectsProvider.Component(module);
            if (component == null)
            {
                return null;
            }

            using (var codeModule = component.CodeModule)
            {
                return GetCurrentLogicalLine(codeModule);
            }
        }

        public Selection GetSelection(QualifiedModuleName module)
        {
            using (var component = _projectsProvider.Component(module))
            using (var codeModule = component.CodeModule)
            using (var pane = codeModule.CodePane)
            {
                return pane.Selection;
            }
        }
    }
}
