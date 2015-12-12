using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace Rubberduck.SmartIndenter
{
    public class Indenter : IIndenter
    {
        private readonly VBE _vbe;
        private readonly IIndenterSettings _settings;

        private readonly Stack<string> _inProcedure = new Stack<string>();
        private readonly Stack<string> _inCode = new Stack<string>();
        private readonly Stack<string> _outProcedure = new Stack<string>();
        private readonly Stack<string> _outCode = new Stack<string>();

        private string[] _declares;
        private string[] _functionAlign;

        public Indenter(VBE vbe, IIndenterSettings settings)
        {
            _vbe = vbe;
            _settings = settings;
        }

        private int _originalTopLine;
        private Selection _originalSelection;

        public event EventHandler<IndenterProgressEventArgs> ReportProgress;

        private void OnReportProgress(string moduleName, int progress, int max)
        {
            var handler = ReportProgress;
            if (handler != null)
            {
                var args = new IndenterProgressEventArgs(moduleName, progress, max);
                handler.Invoke(this, args);
            }
        }

        public void IndentCurrentProcedure()
        {
            var pane = _vbe.ActiveCodePane;
            var selection = GetSelection(pane);

            vbext_ProcKind procKind;
            var procName = pane.CodeModule.get_ProcOfLine(selection.StartLine, out procKind);

            if (string.IsNullOrEmpty(procName))
            {
                procName = null;
            }

            Indent(pane.CodeModule.Parent, procName, selection);
        }

        public void IndentCurrentModule()
        {
            var pane = _vbe.ActiveCodePane;
            Indent(pane.CodeModule.Parent);
        }

        public void Indent(VBProject project)
        {
            if (project == null)
            {
                throw new ArgumentNullException("project");
            }
            if (project.Protection == vbext_ProjectProtection.vbext_pp_locked)
            {
                throw new InvalidOperationException("Project is protected.");
            }
            
            var lineCount = 0; // to set progressbar max value
            if (project.VBComponents.Cast<VBComponent>().All(component => !HasCode(component.CodeModule, ref lineCount)))
            {
                throw new InvalidOperationException("Project contains no code.");
            }

            _originalTopLine = _vbe.ActiveCodePane.TopLine;
            _originalSelection = GetSelection(_vbe.ActiveCodePane);

            var progress = 0; // to set progressbar value
            foreach (var component in project.VBComponents.Cast<VBComponent>().Where(component => HasCode(component.CodeModule)))
            {
                Indent(component, true, progress);
                progress += component.CodeModule.CountOfLines;
            }

            _vbe.ActiveCodePane.TopLine = _originalTopLine;
            _vbe.ActiveCodePane.SetSelection(_originalSelection.StartLine, _originalSelection.StartColumn, _originalSelection.EndLine, _originalSelection.EndColumn);
        }

        private static bool HasCode(CodeModule module, ref int lineCount)
        {
            lineCount += module.CountOfLines;
            for (var i = 0; i < module.CountOfLines; i++)
            {
                if (!string.IsNullOrWhiteSpace(module.get_Lines(i, 1)))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool HasCode(CodeModule module)
        {
            for (var i = 0; i < module.CountOfLines; i++)
            {
                if (!string.IsNullOrWhiteSpace(module.get_Lines(i, 1)))
                {
                    return true;
                }
            }

            return false;
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

        public void Indent(VBComponent module, bool reportProgress = true, int linesAlreadyRebuilt = 0)
        {
            var lineCount = module.CodeModule.CountOfLines;
            if (lineCount == 0)
            {
                return;
            }

            var codeLines = module.CodeModule.get_Lines(1, lineCount).Split('\n');
            Indent(codeLines, module.Name, reportProgress, linesAlreadyRebuilt);
        }

        public void Indent(VBComponent module, string procedureName, Selection selection, bool reportProgress = true, int linesAlreadyRebuilt = 0)
        {
            var lineCount = module.CodeModule.CountOfLines;
            if (lineCount == 0)
            {
                return;
            }

            var codeLines = module.CodeModule.get_Lines(selection.StartLine, selection.LineCount).Split('\n');
            Indent(codeLines, procedureName, reportProgress, linesAlreadyRebuilt);
        }

        public void Indent(string[] codeLines, string moduleName, bool reportProgress = true, int linesAlreadyRebuilt = 0)
        {
            if (_settings.EnableUndo)
            {
                // todo: store undo info
            }

            //var initialized = false;
            var inOnContinuedLine = false;
            var noIndent = false;
            var isInsideIfBlock = false;
            var isInsideComment = false;

            var indentCase = _settings.IndentCase;

            if (_inCode.Any() && _inCode.Peek() != "Select Case" && indentCase)
            {
                _inCode.Push("Select Case");
                _outCode.Push("End Select");
            }
            else if (_inCode.Any() && _inCode.Peek() == "Select Case" && indentCase)
            {
                _inCode.Pop();
                _outCode.Pop();
            }

            var firstProcLine = false;
            var atProcedureStart = false;
            var atFirstDim = false;
            var atFirstCont = true;

            var gap = 0;

            var lineCount = 0;
            var commentStart = 0;
            var start = 0;
            var scan = 0;
            var debugAdjustment = 0;

            var indents = 0;
            var indentNext = 0;
            var ins = 0;
            var outs = 0;
            var functionStart = 0;
            var parameterStart = 0;

            var lineAdjust = 0;
            var alreadyPadded = false;

            for (var line = 0; line < codeLines.Length; line++)
            {
                var numberedLine = -1;
                var originalLine = codeLines[line];
                var currentLine = codeLines[line].Trim();

                // todo: report progress

                // if we're not in a continued line, initialize some variables
                if (!(inOnContinuedLine || isInsideComment))
                {
                    atProcedureStart = false;
                    indentNext = 0;
                    commentStart = 0;
                    indents += debugAdjustment;
                    debugAdjustment = 0;
                    functionStart = 0;
                    parameterStart = 0;

                    // removes explicit line number / replace it with spaces
                    var i = currentLine.IndexOf(' ');
                    if (i > 0 && int.TryParse(currentLine, out numberedLine))
                    {
                        currentLine = currentLine.Substring(i).Trim();
                        originalLine = new string(' ', i) + originalLine.Substring(i);
                    }

                    // is there anything on the line?
                    if (currentLine.Length > 0)
                    {
                        // remove leading tabs, add extra space at the end
                        currentLine = currentLine.TrimStart('\t') + ' ';

                        if (isInsideComment)
                        {
                            // inside a multiline comment - indent to line up the comment text
                            currentLine = new string(' ', commentStart) + currentLine;

                            // remember if we're in a continued comment line
                            isInsideComment = currentLine.EndsWith(" _");

                            goto PTR_REPLACE_LINE;
                        }

                        // remember the position of the line segment
                        start = 1;
                        scan = 0;

                        if (inOnContinuedLine && _settings.AlignContinuations)
                        {
                            if (_settings.AlignIgnoreOps && currentLine.StartsWith(", "))
                            {
                                parameterStart = functionStart - 2;
                            }

                            // todo: test this logic. VB6 logical operator precedence might not match that of C#.
                            if (_settings.AlignIgnoreOps && !currentLine.StartsWith(", ")
                                && (currentLine.Substring(1, 1) == " " || currentLine.StartsWith(":=")))
                            {
                                currentLine = new string(' ', parameterStart - 3) + currentLine;
                                lineAdjust += parameterStart - 3;
                                scan += parameterStart - 3;
                            }
                            else
                            {
                                currentLine = new string(' ', parameterStart - 1) + currentLine;
                                lineAdjust += parameterStart - 1;
                                scan += parameterStart - 1;
                            }

                            alreadyPadded = true;
                        }

                        // scan through the line, char by char, checking for strings, multi-statement lines and comments
                        do
                        {
                            scan++;
                            var item = FindFirstSpecialItemOrDefault(currentLine, ref scan);
                            switch (item)
                            {
                                case "":
                                    //throw new NotImplementedException();
                                    break;

                                case "\"\"":
                                    //throw new NotImplementedException();
                                    break;

                                case ": ":
                                    //throw new NotImplementedException();
                                    break;

                                case " As ":
                                    //throw new NotImplementedException();
                                    break;

                                case "'":
                                case "Rem ":
                                    isInsideComment = true;
                                    //throw new NotImplementedException();
                                    break;

                                case "Stop ":
                                case "Debug.Print ":
                                case "Debug.Assert ":
                                    //throw new NotImplementedException();
                                    break;

                                case "#If ":
                                case "#ElseIf ":
                                case "#Else ":
                                case "#End If":
                                case "#Const ":
                                    //throw new NotImplementedException();
                                    break;
                            }

                        } while (scan <= currentLine.Length);

                        // do we have some code left to check?
                        // i.e. a line without a commnet or the last segment of a multi-statement line

                        if (start < currentLine.Length)
                        {
                            if (!inOnContinuedLine)
                            {
                                atProcedureStart = false;
                            }

                            //CheckLine(currentLine.Substring(start), ins, outs, atProcedureStart);
                            if (atProcedureStart)
                            {
                                atFirstDim = true;
                            }

                            if (start == 1)
                            {
                                indents -= outs;
                                if (indents < 0)
                                {
                                    indents = 0;
                                }

                                indentNext += ins - outs;
                            }
                        }

                        // start from the left at each procedure start
                        if (firstProcLine)
                        {
                            indents = 0;
                        }

                        // line continuations
                        if (inOnContinuedLine)
                        {
                            if (!_settings.AlignContinuations)
                            {
                                currentLine = new string(' ', (indents + 2) * _settings.IndentSpaces) + currentLine;
                            }
                        }
                        else
                        {
                            // check if we start with a declaration item
                            var align = false;
                            if (_settings.IndentProcedure && atFirstDim && !_settings.IndentDim && !atProcedureStart)
                            {
                                if (_declares.Any(declaration => currentLine.StartsWith(declaration + ' ')))
                                {
                                    align = true;
                                }

                                // not a declaration item to left-align, so pad it out
                                if (!align)
                                {
                                    if (!atProcedureStart)
                                    {
                                        atFirstDim = true;
                                    }

                                    currentLine = new string(' ', indents * _settings.IndentSpaces) + currentLine;
                                }
                            }

                            inOnContinuedLine = currentLine.EndsWith(" _");
                        }

                        // anything there?
                    }

                PTR_REPLACE_LINE:
                    // add the coe line number back in
                    if (numberedLine > -1)
                    {
                        var lineNumber = numberedLine.ToString();
                        if (currentLine.Substring(0, lineNumber.Length + 1).Trim().Length == 0)
                        {
                            currentLine = lineNumber + currentLine.Substring(lineNumber.Length + 1);
                        }
                        else
                        {
                            currentLine = lineNumber + currentLine.Trim();
                        }
                    }

                    codeLines[lineCount] = currentLine.TrimEnd();

                    // if it's not a continued line, update the indenting for the following lines
                    if (!inOnContinuedLine)
                    {
                        indents += indentNext;
                        indentNext = 0;
                        if (indents < 0)
                        {
                            indents = 0;
                        }
                    }
                    else
                    {
                        // a continued line, so if we're not in a comment and we want smart continuing,
                        // work out which to continue from
                        if (_settings.AlignContinuations && !isInsideComment)
                        {
                            if (currentLine.Trim().StartsWith("& ") || currentLine.Trim().StartsWith("+ "))
                            {
                                currentLine = "  " + currentLine;
                            }

                            //functionStart = FunctionAlign(currentLine, atFirstCont, parameterStart);
                            if (functionStart == 0)
                            {
                                functionStart = (indents + 2) * _settings.IndentSpaces;
                                parameterStart = functionStart;
                            }
                        }
                    }
                }

                atFirstCont = !inOnContinuedLine;
            }
        }

        private static readonly string[] ProcedureLevelScopeTokens =
        {
            string.Empty, "Public", "Private", "Friend"
        };

        private static readonly string[] ProcedureLevelStaticTokens =
        {
            string.Empty, "Static"
        };

        private static readonly string[] ProcedureLevelTypeTokens =
        {
            "Sub", "Function", "Property Let", "Property Get", "Property Set", "Type", "Enum"
        };

        private static readonly string[] ProcedureLevelOutdentingMatch =
        {
            "End Sub", "End Function", "End Property", "End Type", "End Enum"
        };

        private static readonly string[] InsideProcedureIndentingCompilerStuffMatch =
        {
            "#If", "#ElseIf", "#Else"
        };

        private static readonly string[] InsideProcedureIndentingMatch =
        {
            "If", "ElseIf", "Else", "Select Case", "Case", "With", "For", "Do", "While"
        };

        private static readonly string[] InsideProcedureOutdentingCompilerStuffMatch =
        {
            "#ElseIf", "#Else", "#End If"
        };

        private static readonly string[] InsideProcedureOutdentingMatch =
        {
            "ElseIf", "Else", "End If", "Case", "End Select", "End With", "Next", "Loop", "Wend"
        };

        private static readonly string[] DeclarationLevelMatch =
        {
            "Dim", "Const", "Static", "Public", "Private", "#Const"
        };

        private static readonly string[] InsideCodeLineSpecialHandling =
        {
            "\"\"", ": ", " As ", "'", "Rem ", "Stop ", "Debug.Print ", "Debug.Assert ", "#If ", "#ElseIf ", "#Else ", "#End If", "#Const "
        };

        private static readonly string[] SkipWhenFindingFunctionStart =
        {
            "Set ", "Let ", "LSet ", "RSet ", "Declare Function", "Declare Sub", "Private Declare Function", "Private Declare Sub", "Public Declare Function", "Public Declare Sub"
        };

        private string FindFirstSpecialItemOrDefault(string line, ref int from)
        {
            if (line == null)
            {
                throw new ArgumentNullException("line");
            }

            var first = line.Length;
            var result = string.Empty;

            foreach (var item in InsideCodeLineSpecialHandling.Where(line.Contains))
            {
                var foundAt = line.IndexOf(item, @from, StringComparison.InvariantCulture);
                // is it before any other items?
                if (foundAt > 0 && foundAt < first)
                {
                    first = foundAt;
                    result = item;
                }
            }

            from = first;
            return result;
        }
    }
}
