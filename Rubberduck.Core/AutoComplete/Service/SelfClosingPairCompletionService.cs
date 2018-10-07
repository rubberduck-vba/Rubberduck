using System;
using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing;
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
            if (pair.IsSymetric && input != '\b' &&
                original.Code.Length >= 1 &&
                original.CaretPosition.StartColumn > 0 &&
                original.Code[original.CaretPosition.StartColumn - 1] == pair.ClosingChar)
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
            if (input == Keys.Back)
            {
                return HandleBackspace(pair, original);
            }

            return null;
        }

        private CodeString HandleOpeningChar(SelfClosingPair pair, CodeString original)
        {
            if (pair.IsSymetric)
            {
                // todo
            }

            var nextPosition = original.CaretPosition.ShiftRight();
            var autoCode = new string(new[] { pair.OpeningChar, pair.ClosingChar });
            var lines = original.Lines;
            var line = lines[original.CaretPosition.StartLine];
            lines[original.CaretPosition.StartLine] = string.IsNullOrEmpty(original.Code) 
                    ? autoCode 
                    : original.CaretPosition.StartColumn == line.Length 
                        ? line + autoCode 
                        : line.Insert(original.CaretPosition.StartColumn, autoCode);

            return new CodeString(string.Join("\r\n", lines), nextPosition, new Selection(original.SnippetPosition.StartLine, 1, original.SnippetPosition.EndLine, 1));
        }

        private CodeString HandleClosingChar(SelfClosingPair pair, CodeString original)
        {
            if (pair.IsSymetric)
            {
                return null;
            }

            var isBalanced = original.Code.Count(c => c == pair.OpeningChar) ==
                             original.Code.Count(c => c == pair.ClosingChar);
            var nextIsClosingChar = original.CaretLine.Length > original.CaretCharIndex &&  original.CaretLine[original.CaretCharIndex] == pair.ClosingChar;

            if (isBalanced && nextIsClosingChar)
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
            var previous = Math.Max(0, position.StartColumn - 1);
            var next = Math.Min(line.Length - 1, position.StartColumn);

            var previousChar = line[previous];
            var nextChar = line[next];

            if (original.CaretPosition.EndColumn < next && previousChar == pair.OpeningChar && nextChar == pair.ClosingChar)
            {
                if (line.Length == 2)
                {
                    // entire line consists in the self-closing pair itself
                    return new CodeString(string.Empty, default, Selection.Empty.ShiftRight());
                }
                else
                {
                    lines[original.CaretPosition.StartLine] = line.Remove(previous, 2);
                    return new CodeString(string.Join("\r\n", lines), original.CaretPosition.ShiftLeft(), original.SnippetPosition);
                }
            }

            if (previous < line.Length - 1 && previousChar == pair.OpeningChar)
            {
                Selection closingTokenPosition;
                closingTokenPosition = line[Math.Min(line.Length - 1, next)] == pair.ClosingChar
                    ? position
                    : FindMatchingTokenPosition(pair, original);
                
                if (closingTokenPosition != default)
                {
                    var closingLine = lines[closingTokenPosition.EndLine].Remove(closingTokenPosition.StartColumn, 1);
                    lines[closingTokenPosition.EndLine] = closingLine;

                    if (closingLine == pair.OpeningChar.ToString())
                    {
                        lines[closingTokenPosition.EndLine] = string.Empty;
                    }
                    else
                    {
                        var openingLine = lines[position.StartLine].Remove(position.ShiftLeft().StartColumn, 1);
                        lines[position.StartLine] = openingLine;
                    }
                    lines = lines.Where((x, i) => i <= position.StartLine || !string.IsNullOrWhiteSpace(x)).ToArray();
                    if (lines[lines.Length - 1].EndsWith(" _"))
                    {
                        // logical line can't end with a line continuation token...
                        lines[lines.Length - 1] = lines[lines.Length - 1].TrimEnd(' ', '_');
                    }

                    return new CodeString(string.Join("\r\n", lines), original.CaretPosition.ShiftLeft(),
                        new Selection(original.SnippetPosition.StartLine, 1, original.SnippetPosition.EndLine, 1));
                }
            }

            return null;
        }

        private Selection FindMatchingTokenPosition(SelfClosingPair pair, CodeString original)
        {
            var code = original.Code;
            code = code.EndsWith($"{pair.OpeningChar}{pair.ClosingChar}")
                ? code.Substring(0, code.LastIndexOf(pair.ClosingChar) + 1)
                : code;
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
            visitor.Visit(result.parseTree);
            return visitor.Result;
        }



        private class MatchingTokenVisitor : VBAParserBaseVisitor<Selection>
        {
            private readonly SelfClosingPair _pair;
            private readonly CodeString _code;

            public Selection Result { get; private set; }

            public MatchingTokenVisitor(SelfClosingPair pair, CodeString code)
            {
                _pair = pair;
                _code = code;
            }

            public override Selection VisitLiteralExpr([NotNull] VBAParser.LiteralExprContext context)
            {
                if (context.Start.Text.StartsWith(_pair.OpeningChar.ToString())
                    && context.Start.Text.EndsWith(_pair.ClosingChar.ToString()))
                {
                    if (_code.CaretPosition.StartLine == context.Start.Line - 1
                        && _code.CaretPosition.StartColumn == context.Start.Column + 1)
                    {
                        Result = new Selection(context.Start.Line - 1, context.Stop.Column + context.Stop.Text.Length - 1);
                    }
                }
                var inner = context.GetDescendents<VBAParser.LiteralExprContext>();
                foreach (var item in inner)
                {
                    if (context != item)
                    {
                        var result = Visit(item);
                        if (result != default)
                        {
                            Result = result;
                        }
                    }
                }

                return base.VisitLiteralExpr(context);
            }

            public override Selection VisitIndexExpr([NotNull] VBAParser.IndexExprContext context)
            {
                if (context.LPAREN()?.Symbol.Text[0] == _pair.OpeningChar
                    && context.RPAREN()?.Symbol.Text[0] == _pair.ClosingChar)
                {
                    if (_code.CaretPosition.StartLine == context.LPAREN().Symbol.Line - 1
                        && _code.CaretPosition.StartColumn == context.RPAREN().Symbol.Column)
                    {
                        var token = context.RPAREN().Symbol;
                        Result = new Selection(token.Line - 1, token.Column);
                    }
                }
                var inner = context.GetDescendents<VBAParser.IndexExprContext>();
                foreach (var item in inner)
                {
                    if (context != item)
                    {
                        var result = Visit(item);
                        if (result != default)
                        {
                            Result = result;
                        }
                    }
                }

                return base.VisitIndexExpr(context);
            }

            public override Selection VisitArgList([NotNull] VBAParser.ArgListContext context)
            {
                if (context.Start.Text[0] == _pair.OpeningChar
                    && context.Stop.Text[0] == _pair.ClosingChar)
                {
                    if (_code.CaretPosition.StartLine == context.Start.Line - 1
                        && _code.CaretPosition.StartColumn == context.Start.Column + 1)
                    {
                        var token = context.Stop;
                        Result = new Selection(token.Line - 1, token.Column);
                    }
                }
                var inner = context.GetDescendents<VBAParser.ArgListContext>();
                foreach (var item in inner)
                {
                    if (context != item)
                    {
                        var result = Visit(item);
                        if (result != default)
                        {
                            Result = result;
                        }
                    }
                }

                return base.VisitArgList(context);
            }

            public override Selection VisitParenthesizedExpr([NotNull] VBAParser.ParenthesizedExprContext context)
            {
                if (context.Start.Text[0] == _pair.OpeningChar
                    && context.Stop.Text[0] == _pair.ClosingChar)
                {
                    if (_code.CaretPosition.StartLine == context.Start.Line - 1
                        && _code.CaretPosition.StartColumn == context.Start.Column + 1)
                    {
                        var token = context.Stop;
                        Result = new Selection(token.Line - 1, token.Column);
                    }
                }
                var inner = context.GetDescendents<VBAParser.ParenthesizedExprContext>();
                foreach (var item in inner)
                {
                    if (context != item)
                    {
                        var result = Visit(item);
                        if (result != default)
                        {
                            Result = result;
                        }
                    }
                }

                return base.VisitParenthesizedExpr(context);
            }
        }
    }
}
