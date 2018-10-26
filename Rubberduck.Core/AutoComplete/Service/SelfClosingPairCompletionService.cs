using System;
using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.AutoComplete.Service
{
    public class SelfClosingPairCompletionService
    {
        private readonly IShowIntelliSenseCommand _showIntelliSense;

        public SelfClosingPairCompletionService(IShowIntelliSenseCommand showIntelliSense)
        {
            _showIntelliSense = showIntelliSense;
        }

        public CodeString Execute(SelfClosingPair pair, CodeString original, char input)
        {
            var previousCharIsClosingChar =
                original.CaretPosition.StartColumn > 0 &&
                original.CaretLine[original.CaretPosition.StartColumn - 1] == pair.ClosingChar;
            var nextCharIsClosingChar =
                original.CaretPosition.StartColumn < original.CaretLine.Length &&
                original.CaretLine[original.CaretPosition.StartColumn] == pair.ClosingChar;

            if (pair.IsSymetric && input != '\b' &&
                original.Code.Length >= 1 &&
                previousCharIsClosingChar && !nextCharIsClosingChar
                || original.IsComment || (original.IsInsideStringLiteral && !nextCharIsClosingChar))
            {
                return null;
            }

            if (input == pair.OpeningChar)
            {
                var result = HandleOpeningChar(pair, original);
                return result;
            }

            if (input == pair.ClosingChar)
            {
                return HandleClosingChar(pair, original);
            }

            if (input == '\b')
            {
                return Execute(pair, original, Keys.Back);
            }

            return null;
        }

        public CodeString Execute(SelfClosingPair pair, CodeString original, Keys input)
        {
            if (original.IsComment)
            {
                return null;
            }

            if (input == Keys.Back)
            {
                return HandleBackspace(pair, original);
            }

            return null;
        }

        private CodeString HandleOpeningChar(SelfClosingPair pair, CodeString original)
        {
            var nextPosition = original.CaretPosition.ShiftRight();
            var autoCode = new string(new[] { pair.OpeningChar, pair.ClosingChar });
            var lines = original.Lines;
            var line = lines[original.CaretPosition.StartLine];

            string newCode;
            if (string.IsNullOrEmpty(line))
            {
                newCode = autoCode;
            }
            else if (pair.IsSymetric && original.CaretPosition.StartColumn < line.Length && line[original.CaretPosition.StartColumn] == pair.ClosingChar)
            {
                newCode = line;
            }
            else
            {
                newCode = original.CaretPosition.StartColumn >= line.Length
                    ? line + autoCode
                    : line.Insert(original.CaretPosition.StartColumn, autoCode);
            }
            lines[original.CaretPosition.StartLine] = newCode;

            return new CodeString(string.Join("\r\n", lines), nextPosition, new Selection(original.SnippetPosition.StartLine, 1, original.SnippetPosition.EndLine, 1));
        }

        private CodeString HandleClosingChar(SelfClosingPair pair, CodeString original)
        {
            if (pair.IsSymetric)
            {
                return null;
            }

            var nextIsClosingChar = original.CaretLine.Length > original.CaretCharIndex &&  
                                    original.CaretLine[original.CaretCharIndex] == pair.ClosingChar;
            if (nextIsClosingChar)
            {
                var nextPosition = original.CaretPosition.ShiftRight();
                var newCode = original.Code;

                return new CodeString(newCode, nextPosition, new Selection(original.SnippetPosition.StartLine, 1, original.SnippetPosition.EndLine, 1));
            }
            return null;
        }

        private CodeString HandleBackspace(SelfClosingPair pair, CodeString original)
        {
            return DeleteMatchingTokens(pair, original);
        }

        private CodeString DeleteMatchingTokens(SelfClosingPair pair, CodeString original)
        {
            var position = original.CaretPosition;
            var lines = original.Lines;

            var line = lines[original.CaretPosition.StartLine];
            if (line.Length == 0)
            {
                // nothing to delete at caret position... bail out.
                return null;
            }

            var previous = Math.Max(0, position.StartColumn - 1);
            var next = Math.Min(line.Length - 1, position.StartColumn);

            var previousChar = line[previous];
            var nextChar = line[next];

            if (original.CaretPosition.StartColumn < next && 
                previousChar == pair.OpeningChar && 
                nextChar == pair.ClosingChar)
            {
                if (line.Length == 2)
                {
                    // entire line consists in the self-closing pair itself.
                    return new CodeString(string.Empty, default, Selection.Empty.ShiftRight());
                }

                // simple case; caret is between the opening and closing chars - remove both.
                lines[original.CaretPosition.StartLine] = line.Remove(previous, 2);
                return new CodeString(string.Join("\r\n", lines), original.CaretPosition.ShiftLeft(), original.SnippetPosition);
            }

            if (previous < line.Length - 1 && previousChar == pair.OpeningChar)
            {
                return DeleteMatchingTokensMultiline(pair, original);
            }

            return null;
        }

        private CodeString DeleteMatchingTokensMultiline(SelfClosingPair pair, CodeString original)
        {
            var position = original.CaretPosition;
            var lines = original.Lines;
            var line = lines[original.CaretPosition.StartLine];
            var next = Math.Min(line.Length - 1, position.StartColumn);

            Selection closingTokenPosition;
            closingTokenPosition = line[Math.Min(line.Length - 1, next)] == pair.ClosingChar
                ? position
                : FindMatchingTokenPosition(pair, original);

            if (closingTokenPosition == default)
            {
                // could not locate the closing token... bail out.
                return null;
            }

            var closingLine = lines[closingTokenPosition.EndLine].Remove(closingTokenPosition.StartColumn, 1);
            lines[closingTokenPosition.EndLine] = closingLine;

            if (closingLine == pair.OpeningChar.ToString() || closingLine == pair.OpeningChar + " _" || closingLine == pair.OpeningChar + " & _")
            {
                lines[closingTokenPosition.EndLine] = string.Empty;
            }
            else
            {
                var openingLine = lines[position.StartLine].Remove(position.ShiftLeft().StartColumn, 1);
                lines[position.StartLine] = openingLine;
            }

            var finalCaretPosition = original.CaretPosition.ShiftLeft();

            var lastLine = lines[lines.Length - 1];
            if (string.IsNullOrEmpty(lastLine.Trim()))
            {
                lines = lines.Where((x, i) => i <= position.StartLine || !string.IsNullOrWhiteSpace(x)).ToArray();
                lastLine = lines[lines.Length - 1];

                if (lastLine.EndsWith(" _") && finalCaretPosition.StartLine == lines.Length - 1)
                {
                    finalCaretPosition = HandleBackspaceContinuations(lines, finalCaretPosition);
                }
            }

            var caretLine = lines[finalCaretPosition.StartLine];
            if (caretLine.EndsWith(" _") && finalCaretPosition.StartLine == lines.Length - 1)
            {
                finalCaretPosition = HandleBackspaceContinuations(lines, finalCaretPosition);
            }
            else if (caretLine.EndsWith("& _") || caretLine.EndsWith("&  _"))
            {
                HandleBackspaceContinuations(lines, finalCaretPosition);
            }

            var nonEmptyLines = lines.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
            var lastNonEmptyLine = nonEmptyLines.Length > 0 ? nonEmptyLines[nonEmptyLines.Length - 1] : null;
            if (lastNonEmptyLine != null)
            {
                if (position.StartLine > nonEmptyLines.Length - 1)
                {
                    // caret is on a now-empty line, shift one line up.
                    finalCaretPosition = new Selection(position.StartLine - 1, lastNonEmptyLine.Length - 1);
                }

                if (lastNonEmptyLine.EndsWith(" _"))
                {
                    var newPosition = HandleBackspaceContinuations(nonEmptyLines, new Selection(nonEmptyLines.Length - 1, 1));
                    if (finalCaretPosition.StartLine == nonEmptyLines.Length - 1)
                    {
                        finalCaretPosition = newPosition;
                    }
                }

                lines = nonEmptyLines;
            }


            // remove any dangling empty lines...
            lines = lines.Where((x, i) => i <= position.StartLine || !string.IsNullOrWhiteSpace(x)).ToArray();

            return new CodeString(string.Join("\r\n", lines), finalCaretPosition,
                new Selection(original.SnippetPosition.StartLine, 1, original.SnippetPosition.EndLine, 1));
        }

        private static Selection HandleBackspaceContinuations(string[] nonEmptyLines, Selection finalCaretPosition)
        {
            var lineIndex = Math.Min(finalCaretPosition.StartLine, nonEmptyLines.Length - 1);
            var line = nonEmptyLines[lineIndex];
            if (line.EndsWith(" _") && lineIndex == nonEmptyLines.Length - 1)
            {
                nonEmptyLines[lineIndex] = line.Remove(line.Length - 2);
                line = nonEmptyLines[lineIndex];
            }

            if (lineIndex == nonEmptyLines.Length - 1)
            {
                line = nonEmptyLines[lineIndex];
            }

            if (line.EndsWith("&"))
            {
                // we're not concatenating anything anymore; remove concat operator too.
                var concatOffset = line.EndsWith(" &") ? 2 : 1;
                nonEmptyLines[lineIndex] = line.Remove(line.Length - concatOffset);
            }
            TrimNonEmptyLine(nonEmptyLines, lineIndex, "& vbNewLine");
            TrimNonEmptyLine(nonEmptyLines, lineIndex, "& vbCrLf");
            TrimNonEmptyLine(nonEmptyLines, lineIndex, "& vbCr");
            TrimNonEmptyLine(nonEmptyLines, lineIndex, "& vbLf");

            // we're keeping the closing quote, but let's put the caret inside:
            line = nonEmptyLines[lineIndex];
            var quoteOffset = line.EndsWith("\"") ? 1 : 0;
            finalCaretPosition = new Selection(finalCaretPosition.StartLine, line.Length - quoteOffset);
            return finalCaretPosition;
        }

        private static void TrimNonEmptyLine(string[] nonEmptyLines, int lineIndex, string ending)
        {
            var line = nonEmptyLines[lineIndex];
            if (line.EndsWith(ending, StringComparison.OrdinalIgnoreCase))
            {
                var offset = line.EndsWith(" " + ending, StringComparison.OrdinalIgnoreCase)
                    ? ending.Length + 1
                    : ending.Length;
                nonEmptyLines[lineIndex] = line.Remove(line.Length - offset);
            }
        }

        private Selection FindMatchingTokenPosition(SelfClosingPair pair, CodeString original)
        {
            var code = string.Join("\r\n", original.Lines) + "\r\n";
            code = code.EndsWith($"{pair.OpeningChar}{pair.ClosingChar}")
                ? code.Substring(0, code.LastIndexOf(pair.ClosingChar) + 1)
                : code;

            var leftOfCaret = original.CaretLine.Substring(0, original.CaretPosition.StartColumn + 1);
            var rightOfCaret = original.CaretLine.Substring(original.CaretPosition.StartColumn);

            if (leftOfCaret.Count(c => c == pair.OpeningChar) == 1 &&
                rightOfCaret.Count(c => c == pair.ClosingChar) == 1)
            {
                return new Selection(original.CaretPosition.StartLine,
                    original.CaretLine.LastIndexOf(pair.ClosingChar));
            }

            var result = VBACodeStringParser.Parse(code, p => p.startRule());
            if (((ParserRuleContext)result.parseTree).exception != null)
            {
                result = VBACodeStringParser.Parse(code, p => p.mainBlockStmt());
                if (((ParserRuleContext)result.parseTree).exception != null)
                {
                    result = VBACodeStringParser.Parse(code, p => p.blockStmt());
                    if (((ParserRuleContext)result.parseTree).exception != null)
                    {
                        return default;
                    }
                }
            }
            var visitor = new MatchingTokenVisitor(pair, original);
            var matchingTokenPosition = visitor.Visit(result.parseTree);
            return matchingTokenPosition;
        }



        private class MatchingTokenVisitor : VBAParserBaseVisitor<Selection>
        {
            private readonly SelfClosingPair _pair;
            private readonly CodeString _code;

            public MatchingTokenVisitor(SelfClosingPair pair, CodeString code)
            {
                _pair = pair;
                _code = code;
            }

            protected override bool ShouldVisitNextChild(IRuleNode node, Selection currentResult)
            {
                return currentResult.Equals(default);
            }

            public override Selection VisitLiteralExpr([NotNull] VBAParser.LiteralExprContext context)
            {
                var innerResult = VisitChildren(context);
                if (innerResult != DefaultResult)
                {
                    return innerResult;
                }

                if (context.Start.Text.StartsWith(_pair.OpeningChar.ToString())
                    && context.Start.Text.EndsWith(_pair.ClosingChar.ToString()))
                {
                    if (_code.CaretPosition.StartLine == context.Start.Line - 1
                        && _code.CaretPosition.StartColumn == context.Start.Column + 1)
                    {
                        return new Selection(context.Start.Line - 1, context.Stop.Column + context.Stop.Text.Length - 1);
                    }
                }

                return DefaultResult;
            }

            public override Selection VisitIndexExpr([NotNull] VBAParser.IndexExprContext context)
            {
                var innerResult = VisitChildren(context);
                if (innerResult != DefaultResult)
                {
                    return innerResult;
                }

                if (context.LPAREN()?.Symbol.Text[0] == _pair.OpeningChar
                    && context.RPAREN()?.Symbol.Text[0] == _pair.ClosingChar)
                {
                    if (_code.CaretPosition.StartLine == context.LPAREN().Symbol.Line - 1
                        && _code.CaretPosition.StartColumn == context.LPAREN().Symbol.Column + 1)
                    {
                        var token = context.RPAREN().Symbol;
                        return new Selection(token.Line - 1, token.Column);
                    }
                }

                return DefaultResult;
            }

            public override Selection VisitArgList([NotNull] VBAParser.ArgListContext context)
            {
                var innerResult = VisitChildren(context);
                if (innerResult != DefaultResult)
                {
                    return innerResult;
                }

                if (context.Start.Text[0] == _pair.OpeningChar
                    && context.Stop.Text[0] == _pair.ClosingChar)
                {
                    if (_code.CaretPosition.StartLine == context.Start.Line - 1
                        && _code.CaretPosition.StartColumn == context.Start.Column + 1)
                    {
                        var token = context.Stop;
                        return new Selection(token.Line - 1, token.Column);
                    }
                }

                return DefaultResult;
            }

            public override Selection VisitParenthesizedExpr([NotNull] VBAParser.ParenthesizedExprContext context)
            {
                var innerResult = VisitChildren(context);
                if (innerResult != DefaultResult)
                {
                    return innerResult;
                }

                if (context.Start.Text[0] == _pair.OpeningChar
                    && context.Stop.Text[0] == _pair.ClosingChar)
                {
                    if (_code.CaretPosition.StartLine == context.Start.Line - 1
                        && _code.CaretPosition.StartColumn == context.Start.Column + 1)
                    {
                        var token = context.Stop;
                        return new Selection(token.Line - 1, token.Column);
                    }
                }

                return DefaultResult;
            }
        }
    }
}
