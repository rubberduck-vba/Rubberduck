using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class ConditionalCompilationIfExpression : Expression
    {
        private readonly IExpression _ifCondTokens;
        private readonly IExpression _ifCond;
        private readonly IExpression _ifBlockTokens;
        private readonly IEnumerable<Tuple<IExpression, IExpression, IExpression>> _elseIfCodeCondBlocks;
        private readonly IExpression _elseCondTokens;
        private readonly IExpression _elseBlockTokens;
        private readonly IExpression _endIfTokens;

        public ConditionalCompilationIfExpression(
            IExpression ifCondTokens,
            IExpression ifCond,
            IExpression ifBlockTokens,
            IEnumerable<Tuple<IExpression, IExpression, IExpression>> elseIfCodeCondBlocks,
            IExpression elseCondTokens,
            IExpression elseBlockTokens,
            IExpression endIfTokens)
        {
            _ifCondTokens = ifCondTokens;
            _ifCond = ifCond;
            _ifBlockTokens = ifBlockTokens;
            _elseIfCodeCondBlocks = elseIfCodeCondBlocks;
            _elseCondTokens = elseCondTokens;
            _elseBlockTokens = elseBlockTokens;
            _endIfTokens = endIfTokens;
        }

        public override IValue Evaluate()
        {
            var tokens = new List<IToken>();
            var conditions = new List<bool>();
            tokens.AddRange(
                new LivelinessExpression(
                    new ConstantExpression(new BoolValue(false)),
                    _ifCondTokens)
                    .Evaluate().AsTokens);

            var ifIsAlive = _ifCond.EvaluateCondition();
            conditions.Add(ifIsAlive);
            tokens.AddRange(
                new LivelinessExpression(
                    new ConstantExpression(new BoolValue(ifIsAlive)),
                    _ifBlockTokens)
                    .Evaluate().AsTokens);

            foreach (var elseIf in _elseIfCodeCondBlocks)
            {
                tokens.AddRange(
                   new LivelinessExpression(
                       new ConstantExpression(new BoolValue(false)),
                       elseIf.Item1)
                       .Evaluate().AsTokens);
                var elseIfIsAlive = !ifIsAlive && elseIf.Item2.EvaluateCondition();
                conditions.Add(elseIfIsAlive);
                tokens.AddRange(
                    new LivelinessExpression(
                        new ConstantExpression(new BoolValue(elseIfIsAlive)),
                        elseIf.Item3)
                        .Evaluate().AsTokens);
            }

            if (_elseCondTokens != null)
            {
                tokens.AddRange(
                   new LivelinessExpression(
                       new ConstantExpression(new BoolValue(false)),
                       _elseCondTokens)
                       .Evaluate().AsTokens);
                var elseIsAlive = conditions.All(condition => !condition);
                tokens.AddRange(
                    new LivelinessExpression(
                        new ConstantExpression(new BoolValue(elseIsAlive)),
                        _elseBlockTokens)
                        .Evaluate().AsTokens);
            }
            tokens.AddRange(
                  new LivelinessExpression(
                      new ConstantExpression(new BoolValue(false)),
                      _endIfTokens)
                      .Evaluate().AsTokens);
            return new TokensValue(tokens);
        }
    }
}
