using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;
using System;
using System.Linq;
using System.Windows.Forms;

namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public class SelfClosingPairCompletionService
    {
        private readonly IShowIntelliSenseCommand _showIntelliSense;

        public SelfClosingPairCompletionService(IShowIntelliSenseCommand showIntelliSense)
        {
            _showIntelliSense = showIntelliSense;
        }

        public CodeString Execute(SelfClosingPair pair, CodeString original, char input, ICodeStringPrettifier prettifier = null)
        {
            if (input == pair.OpeningChar)
            {
                var result = HandleOpeningChar(pair, original);
                if (result != default && prettifier != null)
                {
                    if (prettifier.IsSpacingUnchanged(result, original))
                    {
                        //_showIntelliSense?.Execute(); /* lovely VBE makes a loud "DING!!" if the command has no effect */
                        return result;
                    }
                    
                    return default;
                }

                return result;
            }

            if (input == pair.ClosingChar)
            {
                return HandleClosingChar(pair, original);
            }

            return default;
        }

        public CodeString Execute(SelfClosingPair pair, CodeString original, Keys input)
        {
            if (input == Keys.Back)
            {
                return HandleBackspace(pair, original);
            }

            return default;
        }

        private CodeString HandleOpeningChar(SelfClosingPair pair, CodeString original)
        {
            var nextPosition = original.CaretPosition.ShiftRight();
            var autoCode = new string(new[] { pair.OpeningChar, pair.ClosingChar });
            var lines = original.Code.Split('\n');
            var line = lines[original.CaretPosition.StartLine];
            lines[original.CaretPosition.StartLine] = line.Insert(original.CaretPosition.StartColumn, autoCode);

            return new CodeString(string.Join("\n", lines), nextPosition, original.SnippetPosition);
        }

        private CodeString HandleClosingChar(SelfClosingPair pair, CodeString original)
        {
            if (original.Code.Count(c => c == pair.OpeningChar) == original.Code.Count(c => c == pair.ClosingChar))
            {
                var nextPosition = original.CaretPosition.ShiftRight();
                var newCode = original.Code;

                return new CodeString(newCode, nextPosition, original.SnippetPosition);
            }
            return default;
        }

        private CodeString HandleBackspace(SelfClosingPair pair, CodeString original)
        {
            return DeleteMatchingTokens(pair, original);
        }

        private CodeString DeleteMatchingTokens(SelfClosingPair pair, CodeString original)
        {
            var position = original.CaretPosition;
            var lines = original.Lines;

            var previous = Math.Max(0, position.StartColumn - 1);
            var next = previous + 1;

            var line = lines[original.CaretPosition.StartLine];
            if (original.CaretPosition.EndColumn < next && line[previous] == pair.OpeningChar && line[next] == pair.ClosingChar)
            {
                if (line.Length == 2)
                {
                    return new CodeString(string.Empty, default, Selection.Empty.ShiftRight());
                }
                lines[original.CaretPosition.StartLine] = line.Length == 2 ? string.Empty : line.Remove(previous, 2);
                return new CodeString(string.Join("\n", lines), original.CaretPosition.ShiftLeft(), original.SnippetPosition);
            }

            if (previous < line.Length - 1 && line[previous] == pair.OpeningChar)
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
                        lines[original.CaretPosition.StartLine] = string.Empty;
                    }
                    else
                    {
                        var openingLine = lines[original.CaretPosition.StartLine].Remove(original.CaretPosition.ShiftLeft().StartColumn, 1);
                        lines[original.CaretPosition.StartLine] = openingLine;
                    }

                    return new CodeString(string.Join("\n", lines), original.CaretPosition.ShiftLeft(), original.SnippetPosition);
                }
            }

            return default;
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
                    return default;
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
