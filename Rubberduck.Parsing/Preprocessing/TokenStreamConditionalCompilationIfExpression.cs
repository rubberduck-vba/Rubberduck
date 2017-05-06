using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class TokenStreamConditionalCompilationIfExpression : Expression
    {
        private readonly IExpression _ifCondCode;
        private readonly IExpression _ifCondTokens;
        private readonly IExpression _ifCond;
        private readonly IExpression _ifBlock;
        private readonly IExpression _ifBlockTokens;
        private readonly IEnumerable<Tuple<IExpression, IExpression, IExpression, IExpression, IExpression>> _elseIfCodeCondBlocks;
        private readonly IExpression _elseCondCode;
        private readonly IExpression _elseCondTokens;
        private readonly IExpression _elseBlock;
        private readonly IExpression _elseBlockTokens;
        private readonly IExpression _endIfCode;
        private readonly IExpression _endIfTokens;

        public TokenStreamConditionalCompilationIfExpression(
            IExpression ifCondCode,
            IExpression ifCondTokens,
            IExpression ifCond,
            IExpression ifBlock,
            IExpression ifBlockTokens,
            IEnumerable<Tuple<IExpression, IExpression, IExpression, IExpression, IExpression>> elseIfCodeCondBlocks,
            IExpression elseCondCode,
            IExpression elseCondTokens,
            IExpression elseBlock,
            IExpression elseBlockTokens,
            IExpression endIfCode,
            IExpression endIfTokens)
        {
            _ifCondCode = ifCondCode;
            _ifCondTokens = ifCondTokens;
            _ifCond = ifCond;
            _ifBlock = ifBlock;
            _ifBlockTokens = ifBlockTokens;
            _elseIfCodeCondBlocks = elseIfCodeCondBlocks;
            _elseCondCode = elseCondCode;
            _elseCondTokens = elseCondTokens;
            _elseBlock = elseBlock;
            _elseBlockTokens = elseBlockTokens;
            _endIfCode = endIfCode;
            _endIfTokens = endIfTokens;
        }

        public override IValue Evaluate()
        {
            StringBuilder builder = new StringBuilder();
            List<bool> conditions = new List<bool>();
            builder.Append(
                new TokenStreamLivelinessExpression(
                    new ConstantExpression(new BoolValue(false)),
                    _ifCondCode,
                    _ifCondTokens)
                    .Evaluate().AsString);

            var ifIsAlive = _ifCond.EvaluateCondition();
            conditions.Add(ifIsAlive);
            builder.Append(
                new TokenStreamLivelinessExpression(
                    new ConstantExpression(new BoolValue(ifIsAlive)),
                    _ifBlock,
                    _ifBlockTokens)
                    .Evaluate().AsString);

            foreach (var elseIf in _elseIfCodeCondBlocks)
            {
                builder.Append(
                   new TokenStreamLivelinessExpression(
                       new ConstantExpression(new BoolValue(false)),
                       elseIf.Item1,
                       elseIf.Item2)
                       .Evaluate().AsString);
                var elseIfIsAlive = !ifIsAlive && elseIf.Item3.EvaluateCondition();
                conditions.Add(elseIfIsAlive);
                builder.Append(
                    new TokenStreamLivelinessExpression(
                        new ConstantExpression(new BoolValue(elseIfIsAlive)),
                        elseIf.Item4,
                        elseIf.Item5)
                        .Evaluate().AsString);
            }

            if (_elseCondCode != null)
            {
                builder.Append(
                   new TokenStreamLivelinessExpression(
                       new ConstantExpression(new BoolValue(false)),
                       _elseCondCode,
                       _elseCondTokens)
                       .Evaluate().AsString);
                var elseIsAlive = conditions.All(condition => !condition);
                builder.Append(
                    new TokenStreamLivelinessExpression(
                        new ConstantExpression(new BoolValue(elseIsAlive)),
                        _elseBlock,
                        _elseBlockTokens)
                        .Evaluate().AsString);
            }
            builder.Append(
                  new TokenStreamLivelinessExpression(
                      new ConstantExpression(new BoolValue(false)),
                      _endIfCode,
                      _endIfTokens)
                      .Evaluate().AsString);
            return new StringValue(builder.ToString());
        }
    }
}
