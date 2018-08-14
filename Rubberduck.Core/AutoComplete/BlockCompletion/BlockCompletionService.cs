using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Rubberduck.AutoComplete.BlockCompletion
{
    //public class BlockCompletionService
    //{
    //    public BlockCompletionService() { }

    //    public BlockCompletionService(IReadOnlyList<BlockCompletion> blockCompletions)
    //    {
    //        BlockCompletions = blockCompletions;
    //    }

    //    public IReadOnlyList<BlockCompletion> BlockCompletions { get; } = new[]
    //    {
    //        new BlockCompletion("Dim", "dim", "Dim ${1:identifier} As ${2:Variant}$0".Split('\n'), "Dim statement", onlyValidInScope:false),
    //        new BlockCompletion("Private", "priv", "Private ${1:identifier} As ${2:Variant}$0".Split('\n'), "Private declaration statement", onlyValidInScope:false, validInScope:false),
    //        new BlockCompletion("Public", "pub", "Public ${1:identifier} As ${2:Variant}$0".Split('\n'), "Public declaration statement", onlyValidInScope:false, validInScope:false),
    //        new BlockCompletion("IfBlock", "if", "If ${1:condition} Then\n\t$1\nEnd If".Split('\n'), "If...EndIf conditional block", p => p.ifStmt()),
    //        //new BlockCompletion("IfStatement", "if", "If ${1:condition} Then ${1:statement}".Split('\n'), "If...Then conditional statement", p => p.singleLineIfStmt()),
    //        new BlockCompletion("ForEach", "foreach", "For Each ${1:item} In ${2:items}\n\t$0\nNext".Split('\n'), "For Each...Next loop", p => p.forEachStmt()),
    //        new BlockCompletion("For", "for", "For ${1:i} = ${2:lower} To ${3:upper}\n\t$0\nNext".Split('\n'), "For...Next loop", p => p.forNextStmt()),
    //        new BlockCompletion("DoLoop", "do", "Do\n\t$0\nLoop".Split('\n'), "Do...Loop loop", p => p.doLoopStmt()),
    //        new BlockCompletion("DoWhileLoop", "dowl", "Do While ${1:condition}\n\t$0\nLoop".Split('\n'), "Do While...Loop loop", p => p.doLoopStmt()),
    //        new BlockCompletion("DoWhile", "dow", "Do\n\t$0\nWhile ${1:condition}".Split('\n'), "Do...While loop", p => p.doLoopStmt()),
    //        new BlockCompletion("DoUntilLoop", "doul", "Do Until ${1:condition}\n\t$0\nLoop".Split('\n'), "Do Until...Loop loop", p => p.doLoopStmt()),
    //        new BlockCompletion("DoUntil", "dou", "Do\n\t$0\nUntil ${1:condition}".Split('\n'), "Do...Until loop", p => p.doLoopStmt()),
    //        new BlockCompletion("While", "while", "While ${1:condition}\n\t$0\nWend".Split('\n'), "While...Wend loop", p => p.whileWendStmt()),
    //        new BlockCompletion("With", "with", "With ${1:obj}\n\t$0\nEnd With".Split('\n'), "With...End With block", p => p.withStmt()),
    //        new BlockCompletion("Enum", "enum", "Enum ${1:identifier}\n\t$0\nEnd Enum".Split('\n'), "Enum...End Enum block", p => p.enumerationStmt(), onlyValidInScope:false, validInScope:false),
    //        new BlockCompletion("Type", "type", "Type ${1:identifier}\n\t$0\nEnd Type".Split('\n'), "Type...End Type block", p => p.type(), onlyValidInScope:false, validInScope:false),
    //        new BlockCompletion("SubProc", "sub", "Sub ${1:identifier}()\n\t$0\nEnd Sub".Split('\n'), "Sub...End Sub scope", p => p.subStmt()),
    //        new BlockCompletion("FunctionProc", "func", "Function ${1:identifier}() As ${2:Variant}\n\t$0\nEnd Function".Split('\n'), "Function...End Function scope", p => p.functionStmt()),
    //        new BlockCompletion("PropGetProc", "pget", "Public Property Get ${1:identifier}() As ${2:Variant}\n\t$0\nEnd Property".Split('\n'), "Property Get...End Property scope", p => p.propertyGetStmt()),
    //        new BlockCompletion("PropLetProc", "plet", "Public Property Let ${1:identifier}(ByVal ${2:value} As ${3:Variant})\n\t$0\nEnd Property".Split('\n'), "Property Let...End Property scope", p => p.propertyLetStmt()),
    //        new BlockCompletion("PropSetProc", "pset", "Public Property Set ${1:identifier}(ByVal ${2:value} As ${3:Object})\n\t$0\nEnd Property".Split('\n'), "Property Set...End Property scope", p => p.propertySetStmt()),
    //    };

    //    public Keys[] CaptureKeys { get; } = new[] { Keys.Tab, Keys.Enter };
    //    public Keys[] CancelKeys { get; } = new[] { Keys.Escape };

    //    private int _currentPlaceholderIndex;
    //    private string _stdIndent = "    "; // todo: assign from settings in ctor

    //    private BlockCompletion _current;
    //    public BlockCompletion Current
    //    {
    //        get { return _current; }
    //        set // public for testing
    //        {
    //            if (_current != value || value == null)
    //            {
    //                _currentPlaceholderIndex = 0;
    //            }
    //            if (_current != value)
    //            {
    //                _current = value;
    //            }
    //        }
    //    }

    //    public bool IsMatch(string code, out IEnumerable<BlockCompletion> matches)
    //    {
    //        matches = BlockCompletions.Where(completion => (" " + code.Trim()).Equals(" " + completion.Prefix, System.StringComparison.InvariantCultureIgnoreCase));
    //        return matches.Any();
    //    }

    //    /// <summary>
    //    /// Gets the new code and selection given a keypress. Sets service state.
    //    /// </summary>
    //    /// <param name="keypress"></param>
    //    /// <param name="currentLine">The entire current line</param>
    //    /// <param name="pSelection">The 1-based code pane selection</param>
    //    /// <param name="codeModule"></param>
    //    /// <returns></returns>
    //    public (string, Selection) Run(Keys keypress, string currentLine, Selection pSelection, ICodeModule codeModule)
    //    {
    //        var newCode = currentLine;
    //        var newSelection = pSelection;
            
    //        var isCaptureKey = CaptureKeys.Contains(keypress);
    //        var isCancelKey = CancelKeys.Contains(keypress);
    //        var isMatch = IsMatch(currentLine, out IEnumerable<BlockCompletion> matches);
    //        var match = matches.FirstOrDefault();

    //        if (!isCaptureKey && !isCancelKey && !isMatch)
    //        {
    //            return (newCode, newSelection);
    //        }

    //        if (isCancelKey)
    //        {
    //            if (Current != null && currentLine == Current.Prefix)
    //            {
    //                newSelection = new Selection(pSelection.StartLine, pSelection.StartColumn + match.Prefix.Length);
    //            }
    //            Current = null;
    //            return (newCode, newSelection);
    //        }

    //        if (isCaptureKey && isMatch)
    //        {
    //            Current = matches.ToArray()[0];
    //            var indent = string.Concat(currentLine.TakeWhile(c => c == ' '));
    //            newCode = string.Join("\r\n", Current.CodeBody.Select(line => indent + line.Replace("\t", _stdIndent)));
    //            var placeholder = Current.TabStops[_currentPlaceholderIndex];
    //            var offset = placeholder.Position;
    //            if (pSelection.IsSingleCharacter)
    //            {
    //                newSelection = new Selection(pSelection.StartLine + offset.StartLine, offset.StartColumn + 1, pSelection.EndLine + offset.EndLine, pSelection.StartColumn + placeholder.Content.Length + 1);
    //            }
    //            else
    //            {
    //                newSelection = new Selection(pSelection.StartLine + offset.StartLine, pSelection.StartColumn + offset.StartColumn, pSelection.EndLine + offset.EndLine, pSelection.StartColumn + offset.StartColumn + placeholder.Content.Length);
    //            }
    //            return (newCode, newSelection);
    //        }
    //        else if (isMatch && (keypress == Keys.None || keypress == Keys.Space))
    //        {
    //            if (pSelection.IsSingleCharacter)
    //            {
    //                // if next character could trigger block completion, select the matching prefix:
    //                newSelection = new Selection(pSelection.StartLine, pSelection.StartColumn - match.Prefix.Length + 1, pSelection.EndLine, pSelection.EndColumn + 1);
    //            }
    //            else
    //            {
    //                // assume prefix is already selected; set selection to end of line:
    //                newSelection = new Selection(pSelection.StartLine, newCode.Length);
    //            }
    //            return (newCode, newSelection);
    //        }
    //        else if (isCaptureKey && Current != null)
    //        {
    //            string scopeLines;
    //            int startLineOffset;
    //            if (codeModule.CountOfDeclarationLines >= pSelection.StartLine)
    //            {
    //                startLineOffset = 0;
    //                scopeLines = codeModule.GetLines(1, codeModule.CountOfDeclarationLines);
    //            }
    //            else
    //            {
    //                var proc = codeModule.GetProcOfLine(pSelection.StartLine);
    //                var procKind = codeModule.GetProcKindOfLine(pSelection.StartLine);
    //                var procStart = codeModule.GetProcStartLine(proc, procKind);
    //                var procLnCount = codeModule.GetProcCountLines(proc, procKind);
    //                scopeLines = codeModule.GetLines(procStart, procLnCount);
    //                startLineOffset = procStart;
    //            }
    //            var (tree, rewriter) = VBACodeStringParser.Parse(scopeLines, Current.StartRule);

    //            var scopeBlocks = Regex.Matches(scopeLines.Replace("\r\n", "\n"), Current.Syntax).Cast<Match>();
    //            var currentBlock = (
    //                from block in scopeBlocks
    //                select new { Match = block,
    //                             StartLine = startLineOffset + scopeLines.Take(block.Index).Count(e => e == '\n') + 1,
    //                             EndLine = startLineOffset + scopeLines.Take(block.Index + block.Length).Count(e => e == '\n') + 1 }
    //                ).SingleOrDefault();
    //            var tabStops = (from capture in currentBlock.Match.Groups.Cast<Capture>()
    //                           let startLine = startLineOffset + scopeLines.Take(capture.Index).Count(c => c == '\n')
    //                           let startColumn = 1
    //                           select new BlockCompletion.TabStop(capture.Value, new Selection(startLine, startColumn, startLine, startColumn + capture.Value.Length)))
    //                        .Skip(1)
    //                        .ToArray();

    //            var stop = tabStops[_currentPlaceholderIndex];
    //            newSelection = new Selection(stop.Position.StartLine + pSelection.StartLine,
    //                                         stop.Position.StartColumn + pSelection.StartColumn,
    //                                         stop.Position.EndLine + pSelection.EndLine,
    //                                         stop.Position.EndColumn + pSelection.EndColumn);

    //            _currentPlaceholderIndex++;
    //            if (_currentPlaceholderIndex >= Current.TabStops.Count)
    //            {
    //                Current = null;
    //            }
    //        }

    //        return (newCode, newSelection);
    //    }
    //}
}
