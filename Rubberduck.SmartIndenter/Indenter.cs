using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Microsoft.Vbe.Interop;

namespace Rubberduck.SmartIndenter
{
    public class Indenter : IIndenter
    {
        private readonly VBE _vbe;
        private readonly Func<IIndenterSettings> _settings;

        private readonly HashSet<string> _inProcedure = new HashSet<string>();
        private List<string> _inCode = new List<string>();
        private List<string> _outCode = new List<string>();

        public Indenter(VBE vbe, Func<IIndenterSettings> settings)
        {
            _vbe = vbe;
            _settings = settings;

            foreach (var scope in ProcedureLevelScopeTokens)
            {
                foreach (var modifier in ProcedureLevelStaticTokens)
                {
                    foreach (var token in ProcedureLevelTypeTokens)
                    {
                        _inProcedure.Add(string.Join(" ", (new [] {scope, modifier, token}).Where(x => !string.IsNullOrEmpty(x))));
                    }
                }
            }
        }

        private int _originalTopLine;
        private Selection _originalSelection;

        public event EventHandler<IndenterProgressEventArgs> ReportProgress;

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
            var selection = GetSelection(pane);

            vbext_ProcKind procKind;
            var procName = pane.CodeModule.get_ProcOfLine(selection.StartLine, out procKind);

            if (string.IsNullOrEmpty(procName))
            {
                return;
            }

            var startLine = pane.CodeModule.get_ProcStartLine(procName, procKind);
            var endLine = startLine + pane.CodeModule.get_ProcCountLines(procName, procKind);

            selection = new Selection(startLine, 1, endLine, 1);
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

            var codeLines = module.CodeModule.get_Lines(1, lineCount).Replace("\r", string.Empty).Split('\n');
            Indent(codeLines, module.Name, reportProgress, linesAlreadyRebuilt);

            for (var i = 0; i < lineCount; i++)
            {
                if (module.CodeModule.get_Lines(i + 1, 1) != codeLines[i])
                {
                    module.CodeModule.ReplaceLine(i + 1, codeLines[i]);
                }
            }
        }

        public void Indent(VBComponent module, string procedureName, Selection selection, bool reportProgress = true, int linesAlreadyRebuilt = 0)
        {
            var lineCount = module.CodeModule.CountOfLines;
            if (lineCount == 0)
            {
                return;
            }

            var codeLines = module.CodeModule.get_Lines(selection.StartLine, selection.LineCount).Replace("\r", string.Empty).Split('\n');
            Indent(codeLines, procedureName, reportProgress, linesAlreadyRebuilt);

            for (var i = 0; i < selection.EndLine - selection.StartLine; i++)
            {
                if (module.CodeModule.get_Lines(selection.StartLine + i, 1) != codeLines[i])
                {
                    module.CodeModule.ReplaceLine(selection.StartLine + i, codeLines[i]);
                }
            }
        }

        public void Indent(string[] codeLines, string moduleName, bool reportProgress = true, int linesAlreadyRebuilt = 0)
        {
            var settings = _settings.Invoke();

            if (settings.EnableUndo)
            {
                // todo: store undo info
            }

            var inOnContinuedLine = false;
            var noIndent = false;
            var isInsideIfBlock = false;
            var isInsideComment = false;

            var indentCase = settings.IndentCase;

            if (settings.IndentCompilerDirectives)
            {
                _inCode = InsideProcedureIndentingMatch.Union(InsideProcedureIndentingCompilerStuffMatch).ToList();
                _outCode = InsideProcedureOutdentingMatch.Union(InsideProcedureOutdentingCompilerStuffMatch).ToList();
            }
            else
            {
                _inCode = InsideProcedureIndentingMatch.ToList();
                _outCode = InsideProcedureOutdentingMatch.ToList();
            }

            if (_inCode.Any() && _inCode.Last() != "Select Case" && indentCase)
            {
                _inCode.Add("Select Case");
                _outCode.Add("End Select");
            }
            else if (_inCode.Any() && _inCode.Last() == "Select Case" && indentCase)
            {
                _inCode.Remove(_inCode.Last());
                _outCode.Remove(_outCode.Last());
            }

            var atFirstProcLine = false;
            var atProcedureStart = false;
            var atFirstDim = false;
            var atFirstCont = true;

            var gap = 0;
            var commentStart = 0;
            var debugAdjustment = 0;

            var indents = 0;
            var indentNext = 0;
            var functionStart = 0;
            var parameterStart = 0;

            var lineAdjust = 0;

            for (var line = 0; line < codeLines.Length; line++)
            {
                var numberedLine = -1;
                var lineNumber = numberedLine.ToString(CultureInfo.InvariantCulture);
                var originalLine = codeLines[line].Trim();
                var currentLine = codeLines[line].Trim();

                if (reportProgress)
                {
                    // todo: report progress
                }

                // if we're not in a continued line, initialize some variables
                if (!(inOnContinuedLine || isInsideComment))
                {
                    atFirstProcLine = false;
                    indentNext = 0;
                    commentStart = 0;
                    indents += debugAdjustment;
                    debugAdjustment = 0;
                    functionStart = 0;
                    parameterStart = 0;

                    // removes explicit line number / replace it with spaces
                    numberedLine = ResolveLineNumber(ref currentLine, ref originalLine, ref lineNumber);
                }
                // is there anything on the line?
                if (currentLine.Length > 0)
                {                 
                    // remove leading whitespace, add extra space at the end
                    currentLine = currentLine.TrimStart() + ' ';

                    if (isInsideComment)
                    {
                        // inside a multiline comment - indent to line up the comment text
                        currentLine = new string(' ', commentStart) + currentLine;

                        // remember if we're in a continued comment line
                        isInsideComment = currentLine.EndsWith(" _");

                        goto PTR_REPLACE_LINE;
                    }

                    // remember the position of the line segment
                    var start = 1;
                    var scan = 0;

                    if (inOnContinuedLine && settings.AlignContinuations)
                    {
                        if (settings.IgnoreOperatorsInContinuations && currentLine.StartsWith(", "))
                        {
                            parameterStart = functionStart - 2;
                        }

                        // todo: test this logic. VB6 logical operator precedence might not match that of C#.
                        if (settings.IgnoreOperatorsInContinuations && !currentLine.StartsWith(", ")
                            && (currentLine.Substring(1, 1) == " " || currentLine.StartsWith(":=")))
                        {
                            currentLine = new string(' ', parameterStart - 3) + currentLine;
                            lineAdjust = lineAdjust + parameterStart - 3;
                            scan = scan + parameterStart - 3;
                        }
                        else
                        {
                            currentLine = new string(' ', parameterStart - 1) + currentLine;
                            lineAdjust = lineAdjust + parameterStart - 1;
                            scan = scan + parameterStart - 1;
                        }
                    }

                    // scan through the line, char by char, checking for strings, multi-statement lines and comments
                    int outs;
                    int ins;
                    do
                    {
                        var item = FindFirstSpecialItemOrDefault(currentLine, ref scan);
                        switch (item)
                        {
                            case "\"":
                                // start of a string => jump to the end of it
                                scan = currentLine.IndexOf("\"", scan + 1, StringComparison.InvariantCulture);
                                break;

                            case ": ":
                                // a multi-statement line separator => tidy up and continue
                                if (!currentLine.Substring(0, scan + 1).EndsWith("Then:"))
                                {
                                    currentLine = currentLine.Substring(0, scan + 1) + currentLine.Substring(scan + 2);
                                    // check the indenting for the line segment
                                    CheckLine(settings, currentLine, ref noIndent, out ins, out outs, ref atProcedureStart, ref atFirstProcLine, ref isInsideIfBlock);
                                    if (atProcedureStart)
                                    {
                                        atFirstDim = true;
                                    }
                                    indentNext += ins;
                                    if (start == 1)
                                    {
                                        indents = Math.Max(indents - outs, 0);
                                    }
                                    else
                                    {
                                        indentNext -= outs;
                                    }
                                }
                                start = scan + 2;
                                break;

                            case " As ":

                                if (settings.AlignDims)
                                {
                                    var align = noIndent;
                                    if (!align)
                                    {
                                        align = DeclarationLevelMatch.Any(declare => currentLine.Substring(0, declare.Length) == declare);
                                    }

                                    if (align)
                                    {
                                        if (!currentLine.Substring(scan + 2).Contains(" As "))
                                        {
                                            //If mbIndentProc And bFirstDim And Not mbIndentDim And Not mbNoIndent Then
                                            if (!noIndent && atFirstDim && settings.IndentEntireProcedureBody &&
                                                settings.IndentFirstDeclarationBlock)
                                            {
                                                gap = settings.AlignDimColumn - currentLine.Substring(0, scan).TrimEnd().Length;

                                                // adjust for a line number at the start of the line:
                                                if (numberedLine > -1 && lineNumber.Length >= indents*settings.IndentSpaces)
                                                {
                                                    gap -= lineNumber.Length - indents*settings.IndentSpaces - 1;
                                                }
                                            }
                                            gap = Math.Min(gap, 1);
                                        }
                                        else
                                        {
                                            // multiple declarations on the line; don't space out
                                            gap = 1;
                                        }

                                        // work out the new spacing
                                        var left = currentLine.Substring(0, scan).TrimEnd();
                                        currentLine = left + new string(' ', gap) + currentLine.Substring(scan);
                                        scan = left.Length + gap + 3;
                                    }
                                }
                                else
                                {
                                    // not aligning Dims; remove all whitespace
                                    scan = currentLine.Substring(0, scan).TrimEnd().Length;
                                    currentLine = currentLine.Substring(0, scan).TrimEnd() + " " + currentLine.Substring(scan).Trim();
                                    scan += 3;
                                }

                                break;

                            case "'":
                            case "Rem ":
                                // start of a comment: handle end-of-line properly
                                if (scan == 1)
                                {
                                    if (noIndent || atProcedureStart || settings.IndentFirstCommentBlock)
                                    {
                                        // inside procedure: indent to align with code
                                        currentLine = new string(' ', indents*settings.IndentSpaces) + currentLine;
                                        commentStart = scan + settings.IndentSpaces*indents;
                                    }
                                    else if (!atProcedureStart && indents > 0 && settings.IndentEntireProcedureBody)
                                    {
                                        // at the top of the procedure, so indent once if required
                                        currentLine = new string(' ', settings.IndentSpaces) + currentLine;
                                        commentStart = scan + settings.IndentSpaces;
                                    }
                                }
                                else
                                {
                                    if (item == "Rem " && currentLine.Substring(scan - 1, 1) != " " && currentLine.Substring(scan - 1, 1) != ":")
                                    {
                                        break;
                                    }

                                    CheckLine(settings, currentLine.Substring(start, scan - 1), ref noIndent, out ins,out outs, ref atProcedureStart, ref atFirstProcLine, ref isInsideIfBlock);
                                    if (atProcedureStart)
                                    {
                                        atFirstDim = true;
                                    }
                                    indentNext += ins;
                                    if (start == 1)
                                    {
                                        indents = Math.Max(indents - outs, 0);
                                    }
                                    else
                                    {
                                        indentNext -= outs;
                                    }

                                    // get the text before the comment, and the comment text
                                    var left = currentLine.Substring(0, scan - 1);
                                    var right = currentLine.Substring(scan);

                                    // indent the code part of the line
                                    if (inOnContinuedLine && settings.AlignContinuations)
                                    {
                                        currentLine = currentLine.Substring(0, scan - 1).TrimEnd();
                                    }
                                    else
                                    {
                                        if (inOnContinuedLine)
                                        {
                                            currentLine = new string(' ', (indents + 2) * settings.IndentSpaces) + left;
                                        }
                                        else
                                        {
                                            if (atFirstDim && settings.IndentEntireProcedureBody && settings.AlignDims)
                                            {
                                                currentLine = left;
                                            }
                                            else
                                            {
                                                currentLine = new string(' ', indents*settings.IndentSpaces) + left;
                                            }
                                        }
                                    }
                                    inOnContinuedLine = currentLine.Trim().EndsWith(" _");

                                    switch (settings.EndOfLineCommentStyle)
                                    {
                                        case EndOfLineCommentStyle.Absolute:
                                            scan = scan - lineAdjust + originalLine.Length - originalLine.Trim().Length;
                                            gap = scan - currentLine.Length - 1;
                                            break;
                                        case EndOfLineCommentStyle.SameGap:
                                            scan = scan - lineAdjust + originalLine.Length - originalLine.Trim().Length;
                                            gap = scan - originalLine.Substring(0, scan - 1).TrimEnd().Length - 1;
                                            break;
                                        case EndOfLineCommentStyle.StandardGap:
                                            gap = settings.IndentSpaces*2;
                                            break;
                                        case EndOfLineCommentStyle.AlignInColumn:
                                            gap = settings.EndOfLineCommentColumnSpaceAlignment - currentLine.Length - 1;
                                            break;
                                        default:
                                            throw new ArgumentOutOfRangeException();
                                    }

                                    // adjust for a numbered line
                                    if (numberedLine > -1 && lineNumber.Length >= indents*settings.IndentSpaces &&
                                        (settings.EndOfLineCommentStyle == EndOfLineCommentStyle.Absolute ||
                                         settings.EndOfLineCommentStyle == EndOfLineCommentStyle.AlignInColumn))
                                    {
                                        gap -= lineNumber.Length - indents*settings.IndentSpaces - 1;
                                    }
                                    if (gap < 2)
                                    {
                                        gap = settings.IndentSpaces;
                                    }

                                    commentStart = currentLine.Length + gap;
                                    currentLine += new string(' ', gap) + right;
                                }

                                // work out where the text of the comment starts, to align the next line
                                if (currentLine.Substring(commentStart, 4) == "Rem ")
                                {
                                    commentStart += 3;
                                }
                                if (currentLine.Substring(commentStart, 1) == "'")
                                {
                                    commentStart += 1;
                                }
                                while (currentLine.Substring(commentStart, 1) != " ")
                                {
                                    commentStart += 1;
                                }
                                commentStart -= 1;

                                // adjust for a line number at the start of the line
                                if (numberedLine > -1 && lineNumber.Length >= indents*settings.IndentSpaces)
                                {
                                    commentStart += lineNumber.Length - indents*settings.IndentSpaces + 1;
                                }

                                isInsideComment = currentLine.Trim().EndsWith(" _");
                                goto PTR_REPLACE_LINE;
                            case "Stop ":
                            case "Debug.Print ":
                            case "Debug.Assert ":
                                if (start == 1 && scan == 1 && settings.ForceDebugStatementsInColumn1)
                                {
                                    // note: original code seems to subtract the length of originalLine implicitly converted to a string, and trimmed
                                    lineAdjust -= (originalLine.Length - originalLine.TrimStart().Length);
                                    debugAdjustment = indents;
                                    indents = 0;
                                }
                                break;

                            case "#If ":
                            case "#ElseIf ":
                            case "#Else ":
                            case "#End If":
                            case "#Const ":
                                if (start == 1 && scan == 1 && settings.ForceCompilerDirectivesInColumn1)
                                {
                                    // note: original code seems to subtract the length of originalLine implicitly converted to a string, and trimmed
                                    lineAdjust -= (originalLine.Length - originalLine.TrimStart().Length);
                                    debugAdjustment = indents;
                                    indents = 0;
                                }
                                break;
                        }

                        scan++;
                    } while (scan <= currentLine.Length);
                    
                    // do we have some code left to check?
                    // i.e. a line without a commnet or the last segment of a multi-statement line
                    if (start < currentLine.Length)
                    {
                        if (!inOnContinuedLine)
                        {
                            atProcedureStart = false;
                        }

                        CheckLine(settings, currentLine.Substring(start - 1, scan - 1), ref noIndent, out ins, out outs, ref atProcedureStart, ref atFirstProcLine, ref isInsideIfBlock);
                        if (atProcedureStart)
                        {
                            atFirstDim = true;
                        }
                        indentNext += ins;
                        if (start == 1)
                        {
                            indents = Math.Max(indents - outs, 0);
                        }
                        else
                        {
                            indentNext -= outs;
                        }
                    }

                    // start from the left at each procedure start
                    if (atFirstProcLine)
                    {
                        indents = 0;
                    }

                    // line continuations
                    if (inOnContinuedLine)
                    {
                        if (!settings.AlignContinuations)
                        {
                            currentLine = new string(' ', (indents + 2)*settings.IndentSpaces) + currentLine;
                        }
                    }
                    else
                    {
                        // check if we start with a declaration item
                        var align = false;
                        if (settings.IndentEntireProcedureBody && atFirstDim && !settings.IndentFirstDeclarationBlock && !atProcedureStart)
                        {
                            if (DeclarationLevelMatch.Any(declaration => currentLine.StartsWith(declaration + ' ')))
                            {
                                align = true;
                            }
                        }

                        // not a declaration item to left-align, so pad it out
                        if (!align)
                        {
                            if (!atProcedureStart)
                            {
                                atFirstDim = false;
                            }
                            currentLine = new string(' ', indents*settings.IndentSpaces) + currentLine;
                        }                        
                    }
                    inOnContinuedLine = currentLine.EndsWith(" _");               
                    // anything there?
            PTR_REPLACE_LINE:
                    // add the code line number back in
                    if (numberedLine > -1)
                    {
                        if (currentLine.Substring(0, lineNumber.Length + 1).Trim().Length == 0)
                        {
                            currentLine = lineNumber + currentLine.Substring(lineNumber.Length + 1);
                        }
                        else
                        {
                            currentLine = lineNumber + currentLine.Trim();
                        }
                    }

                    codeLines[line] = currentLine.TrimEnd();

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
                        if (settings.AlignContinuations && !isInsideComment)
                        {
                            if (currentLine.Trim().StartsWith("& ") || currentLine.Trim().StartsWith("+ "))
                            {
                                currentLine = "  " + currentLine;
                            }

                            functionStart = FunctionAlign(settings, currentLine, atFirstCont, out parameterStart);
                            if (functionStart == 0)
                            {
                                functionStart = (indents + 2) * settings.IndentSpaces;
                                parameterStart = functionStart;
                            }
                        }
                    }
                }
                atFirstCont = !inOnContinuedLine;
            }
        }

        private static int ResolveLineNumber(ref string currentLine, ref string originalLine, ref string lineNumber)
        {
            int numberedLine;
            var i = currentLine.IndexOf(' ');
            if (i > 0 && int.TryParse(currentLine, out numberedLine))
            {
                currentLine = currentLine.Substring(i).Trim();
                originalLine = new string(' ', i) + originalLine.Substring(i);
                lineNumber = numberedLine.ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                numberedLine = -1;
            }
            return numberedLine;
        }

        private void CheckLine(IIndenterSettings settings, string code, ref bool noIndent, out int ins, out int outs, ref bool atProcedureStart, ref bool atFirstProcLine, ref bool insideIf)
        {

            ins = 0;
            outs = 0;
            var line = code.Trim() + " ";
            if (!noIndent)
            {
                ins += _inCode.Count(value => line.StartsWith(value) && (line.Substring(value.Length, 1) == " " || line.Substring(value.Length, 1) == ":"));
                outs += _outCode.Count(value => line.StartsWith(value) && (line.Substring(value.Length, 1) == " " || (line.Substring(value.Length, 1) == ":" && line.Substring(value.Length + 1, 1) != "=")));
            }

            
            foreach (var value in _inProcedure.Where(value => line.StartsWith(value) && (line.Substring(value.Length, 1) == " " || (line.Substring(value.Length, 1) == ":" && line.Substring(value.Length + 1, 1) != "="))))
            {
                atProcedureStart = true;
                atFirstProcLine = true;

                // don't indent within type or enum constructs
                if (value.EndsWith("Type") || value.EndsWith("Enum"))
                {
                    ins++;
                    noIndent = true;
                }
                else if (!noIndent && settings.IndentEntireProcedureBody)
                {
                    ins++;
                }
            }

            var outMatches = ProcedureLevelOutdentingMatch.Where(value => line.StartsWith(value) && 
                                                                 (line.Substring(value.Length, 1) == " " || 
                                                                  (line.Substring(value.Length, 1) == ":" && 
                                                                   line.Substring(value.Length + 1, 1) != "=")
                                                                 )).Count(value => !value.EndsWith("Type ") && !value.EndsWith("Enum"));

            outs += outMatches;
            if (outMatches > 0)
            {
                
            }
            // special-case handle 'If'; if 'Then' is followed by anything other than a comment, we don't indent.
            if (noIndent || (!insideIf && !code.StartsWith("If ") && !code.StartsWith("#If ")))
            {
                return;
            }

            if (insideIf)
            {
                ins = 1;
            }

            // strip strings
            var i = code.IndexOf('"');
            while (i >= 0)
            {
                var j = code.IndexOf('"', i + 1);
                if (j == -1)
                {
                    j = code.Length;
                }
                code = code.Substring(0, i - 1) + code.Substring(j + 1);
                i = code.IndexOf('"');
            }

            // strip comments
            i = code.IndexOf('\'');
            if (i >= 1)
            {
                code = code.Substring(0, i - 1);
            }

            // allow lines continuations inside the 'If' 
            insideIf = code.Trim().EndsWith(" _");

            // if we have a 'Then' in the line, adding space before & after
            // enables testing for the 'Then' being both within or at the end of the line.
            code = ' ' + code + ' ';
            i = code.IndexOf(" Then ", StringComparison.InvariantCulture);

            if (i >= 0)
            {
                if (code.Substring(i + 5).Trim() != string.Empty)
                {
                    // there's something after the 'Then', we don't indent the 'If'
                    ins = 0;
                }
                // no need to check next time around
                insideIf = false;
            }
        }

        private static readonly Stack<Tuple<string, int>> CurrentAlignment = new Stack<Tuple<string, int>>();

        private int FunctionAlign(IIndenterSettings settings, string line, bool firstLine, out int paramOffset)
        {
            if (firstLine)
            {
                CurrentAlignment.Clear();
            }

            //Convert and numbers at the start of the line to spaces
            int testToken;
            var leftPadding = 0;
            var space = line.IndexOf(' ');
            if (space > 0 && int.TryParse(line, out testToken))
            {
                line = line.Substring(space);
                leftPadding = space + 1;
            }

            leftPadding += (line.Length - line.TrimStart().Length);
            var iFirstThisLine = CurrentAlignment.Count;

            line = line.Trim();
            //Skip over stuff that we don't want to locate the start off
            var skip = SkipWhenFindingFunctionStart.Where(sMatch => line.StartsWith(sMatch)).Sum(sMatch => sMatch.Length) + 2;

            for (var charIndex = skip; charIndex <= line.Length; charIndex++)
            {
                var character = line.Substring(charIndex - 1, 1);
                switch (character)
                {                    
                    case "\"":
                        //A String => jump to the end of it
                        charIndex = line.IndexOf("\"", charIndex + 1, StringComparison.InvariantCulture);
                        break;
                    case "(":
                        //Start of another function => remember this position
                        CurrentAlignment.Push(new Tuple<string, int>("(", charIndex + leftPadding + 2));
                        CurrentAlignment.Push(new Tuple<string, int>(",", charIndex + leftPadding + 3));
                        break;
                    case ")":
                        //Function finished => Remove back to the previous open bracket
                        while (CurrentAlignment.Any() && (!CurrentAlignment.Peek().Item1.Equals("(") || CurrentAlignment.Count == iFirstThisLine))
                        {
                            CurrentAlignment.Pop();
                        }
                        break;
                    case " ":                        
                        if (charIndex + 3 < line.Length && line.Substring(charIndex - 1, 3).Equals(" = "))
                        {
                            //Space before an = sign => remember it to align to later
                            if (!CurrentAlignment.Any(align => align.Item1.Equals("=") || align.Item1.Equals(" ")))
                            {
                                CurrentAlignment.Push(new Tuple<string, int>("=", charIndex + leftPadding + 2));
                            }
                        }
                        else if (!CurrentAlignment.Any() && charIndex < line.Length - 2)
                        {
                            //Space after a name before the end of the line => remember it for later
                            CurrentAlignment.Push(new Tuple<string, int>(" ", charIndex + leftPadding));
                        }
                        else if (charIndex > 5 && line.Substring(charIndex - 5, 6).Equals(" Then "))
                        {
                            //Clear the collection if we find a Then in an If...Then and set the
                            //indenting to align with the bit after the "If "
                            while (CurrentAlignment.Count > 1)
                            {
                                CurrentAlignment.Pop();
                            }
                        }
                        break;
                    case ",":
                        //Start of a new parameter => remember it to align to
                        CurrentAlignment.Push(new Tuple<string, int>(",", charIndex + leftPadding + 2));
                        break;
                    case ":":
                        if (line.Substring(charIndex, 2).Equals(":="))
                        {
                            //A named paremeter => remember to align to after the name
                            CurrentAlignment.Push(new Tuple<string, int>(",", charIndex + leftPadding + 2));
                        }
                        else if (line.Substring(charIndex, 2).Equals(": "))
                        {
                            //A new line section, so clear the brackets
                            CurrentAlignment.Clear();
                            charIndex++;
                        }
                        break;
                }
            }
            //If we end with a comma or a named parameter, get rid of all other comma alignments
            if (line.Substring(line.Length - 3).Equals(", _") || line.Substring(line.Length - 3).Equals(", _"))
            {
                while (CurrentAlignment.Any() && CurrentAlignment.Peek().Item1.Equals(","))
                {                    
                    CurrentAlignment.Pop();
                }
            }

            //If we end with a "( _", remove it and the space alignment after it
            if (line.Substring(line.Length - 3).Equals("( _"))
            {
                CurrentAlignment.Pop(); 
                CurrentAlignment.Pop();
            }

            paramOffset = 0;
            //Get the position of the unmatched bracket and align to that
            foreach (var align in CurrentAlignment)
            {
                if (align.Item1.Equals(","))
                {
                    paramOffset = align.Item2;
                }
                else if (align.Item1.Equals("("))
                {
                    paramOffset = align.Item2 + 1;
                }
                else
                {
                    skip = align.Item2;
                }
            }

            if (skip == 1 || skip >= line.Length + leftPadding - 1)
            {
                if (!CurrentAlignment.Any() && firstLine)
                {
                    skip = settings.IndentSpaces * 2 + leftPadding;
                }
                else
                {
                    skip = leftPadding;
                }
            }

            if (paramOffset == 0)
            {
                paramOffset = skip + 1;
            }
            return skip + 1;
        }

        private static readonly HashSet<string> ProcedureLevelScopeTokens = new HashSet<string>
        {
            string.Empty, "Public", "Private", "Friend"
        };

        private static readonly HashSet<string> ProcedureLevelStaticTokens = new HashSet<string>
        {
            string.Empty, "Static"
        };

        private static readonly HashSet<string> ProcedureLevelTypeTokens = new HashSet<string>
        {
            "Sub", "Function", "Property Let", "Property Get", "Property Set", "Type", "Enum"
        };

        private static readonly HashSet<string> ProcedureLevelOutdentingMatch = new HashSet<string>
        {
            "End Sub", "End Function", "End Property", "End Type", "End Enum"
        };

        private static readonly HashSet<string> InsideProcedureIndentingCompilerStuffMatch = new HashSet<string>
        {
            "#If", "#ElseIf", "#Else"
        };

        private static readonly HashSet<string> InsideProcedureIndentingMatch = new HashSet<string>
        {
            "If", "ElseIf", "Else", "Select Case", "Case", "With", "For", "Do", "While"
        };

        private static readonly HashSet<string> InsideProcedureOutdentingCompilerStuffMatch = new HashSet<string>
        {
            "#ElseIf", "#Else", "#End If"
        };

        private static readonly HashSet<string> InsideProcedureOutdentingMatch = new HashSet<string>
        {
            "ElseIf", "Else", "End If", "Case", "End Select", "End With", "Next", "Loop", "Wend"
        };

        private static readonly HashSet<string> DeclarationLevelMatch = new HashSet<string>
        {
            "Dim", "Const", "Static", "Public", "Private", "#Const"
        };

        private static readonly HashSet<string> InsideCodeLineSpecialHandling = new HashSet<string>
        {
            "\"", ": ", " As ", "'", "Rem ", "Stop ", "Debug.Print ", "Debug.Assert ", "#If ", "#ElseIf ", "#Else ", "#End If", "#Const "
        };

        private static readonly HashSet<string> SkipWhenFindingFunctionStart = new HashSet<string>
        {
            "Set ", "Let ", "LSet ", "RSet ", "Declare Function", "Declare Sub", "Private Declare Function", "Private Declare Sub", "Public Declare Function", "Public Declare Sub"
        };

        private static string FindFirstSpecialItemOrDefault(string line, ref int from)
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
