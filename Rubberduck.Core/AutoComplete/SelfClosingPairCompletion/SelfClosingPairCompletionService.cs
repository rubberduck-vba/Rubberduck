using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Windows.Forms;

namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public class SelfClosingPairCompletionService
    {
        public CodeString Execute(SelfClosingPair pair, CodeString original, char input)
        {
            if (input == pair.OpeningChar)
            {
                return HandleOpeningChar(pair, original);
            }
            else if (input == pair.ClosingChar)
            {
                return HandleClosingChar(pair, original);
            }
            else
            {
                return default;
            }
        }

        public CodeString Execute(SelfClosingPair pair, CodeString original, Keys input)
        {
            if (input == Keys.Back)
            {
                return HandleBackspace(pair, original);
            }
            else
            {
                return default;
            }
        }

        private CodeString HandleOpeningChar(SelfClosingPair pair, CodeString original)
        {
            var nextPosition = original.CaretPosition.ShiftRight();
            var autoCode = new string(new[] { pair.OpeningChar, pair.ClosingChar });
            return new CodeString(original.Code.Insert(original.CaretPosition.StartColumn, autoCode), nextPosition);
        }

        private CodeString HandleClosingChar(SelfClosingPair pair, CodeString original)
        {
            var nextPosition = original.CaretPosition.ShiftRight();
            var newCode = original.Code;

            return new CodeString(newCode, nextPosition);
        }

        private CodeString HandleBackspace(SelfClosingPair pair, CodeString original)
        {
            var lines = original.Lines;
            var line = lines[original.CaretPosition.StartLine];

            return DeleteMatchingTokens(pair, original);
        }

        private CodeString DeleteMatchingTokens(SelfClosingPair pair, CodeString original)
        {
            var position = original.CaretPosition;
            var lines = original.Lines;

            var previous = Math.Max(0, position.StartColumn - 1);
            if (lines[original.CaretPosition.StartLine][previous] == pair.OpeningChar)
            {
                var closingTokenPosition = FindMatchingTokenPosition(pair, original);
                if (closingTokenPosition != default)
                {
                    var closingLine = lines[closingTokenPosition.EndLine].Remove(closingTokenPosition.StartColumn, 1);
                    lines[closingTokenPosition.EndLine] = closingLine;

                    var openingLine = lines[position.StartLine].Remove(original.CaretPosition.ShiftLeft().StartColumn, 1);
                    lines[original.CaretPosition.StartLine] = openingLine;

                    return new CodeString(string.Join("\n", lines), original.CaretPosition.ShiftLeft());
                }

                return default;
            }
            else
            {
                return default;
            }
        }

        private Selection FindMatchingTokenPosition(SelfClosingPair pair, CodeString original)
        {
            var result = VBACodeStringParser.Parse(original, p => p.blockStmt());
            var visitor = new MatchingTokenVisitor(pair, original);
            return visitor.Visit(result.parseTree);
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

            public override Selection VisitLiteralExpr([NotNull] VBAParser.LiteralExprContext context)
            {
                if (context.Start.Text.StartsWith(_pair.OpeningChar.ToString())
                    && context.Start.Text.EndsWith(_pair.ClosingChar.ToString()))
                {
                    if (_code.CaretPosition.StartLine == context.Start.Line - 1
                        && _code.CaretPosition.StartColumn == context.Start.Column + 1)
                    {
                        return new Selection(context.Start.Line - 1, context.Stop.Column + context.Stop.Text.Length - 1);
                    }
                }
                var inner = context.GetDescendents<VBAParser.LiteralExprContext>();
                foreach (var item in inner)
                {
                    if (context != item)
                    {
                        var result = Visit(item);
                        if (result != null)
                        {
                            return result;
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
                        return new Selection(token.Line - 1, token.Column);
                    }
                }
                var inner = context.GetDescendents<VBAParser.IndexExprContext>();
                foreach (var item in inner)
                {
                    if (context != item)
                    {
                        var result = Visit(item);
                        if (result != null)
                        {
                            return result;
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
                        return new Selection(token.Line - 1, token.Column);
                    }
                }
                var inner = context.GetDescendents<VBAParser.ArgListContext>();
                foreach (var item in inner)
                {
                    if (context != item)
                    {
                        var result = Visit(item);
                        if (result != null)
                        {
                            return result;
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
                        return new Selection(token.Line - 1, token.Column);
                    }
                }
                var inner = context.GetDescendents<VBAParser.ParenthesizedExprContext>();
                foreach (var item in inner)
                {
                    if (context != item)
                    {
                        var result = Visit(item);
                        if (result != null)
                        {
                            return result;
                        }
                    }
                }

                return base.VisitParenthesizedExpr(context);
            }
        }
    }
}
