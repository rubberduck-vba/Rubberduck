using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class ConditionalCompilationIfExpression : Expression
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

        public ConditionalCompilationIfExpression(
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
            var tokens = new List<IToken>();
            var conditions = new List<bool>();
            tokens.AddRange(
                new LivelinessExpression(
                    new ConstantExpression(new BoolValue(false)),
                    _ifCondCode,
                    _ifCondTokens)
                    .Evaluate().AsTokens);

            var ifIsAlive = _ifCond.EvaluateCondition();
            conditions.Add(ifIsAlive);
            tokens.AddRange(
                new LivelinessExpression(
                    new ConstantExpression(new BoolValue(ifIsAlive)),
                    _ifBlock,
                    _ifBlockTokens)
                    .Evaluate().AsTokens);

            foreach (var elseIf in _elseIfCodeCondBlocks)
            {
                tokens.AddRange(
                   new LivelinessExpression(
                       new ConstantExpression(new BoolValue(false)),
                       elseIf.Item1,
                       elseIf.Item2)
                       .Evaluate().AsTokens);
                var elseIfIsAlive = !ifIsAlive && elseIf.Item3.EvaluateCondition();
                conditions.Add(elseIfIsAlive);
                tokens.AddRange(
                    new LivelinessExpression(
                        new ConstantExpression(new BoolValue(elseIfIsAlive)),
                        elseIf.Item4,
                        elseIf.Item5)
                        .Evaluate().AsTokens);
            }

            if (_elseCondCode != null)
            {
                tokens.AddRange(
                   new LivelinessExpression(
                       new ConstantExpression(new BoolValue(false)),
                       _elseCondCode,
                       _elseCondTokens)
                       .Evaluate().AsTokens);
                var elseIsAlive = conditions.All(condition => !condition);
                tokens.AddRange(
                    new LivelinessExpression(
                        new ConstantExpression(new BoolValue(elseIsAlive)),
                        _elseBlock,
                        _elseBlockTokens)
                        .Evaluate().AsTokens);
            }
            tokens.AddRange(
                  new LivelinessExpression(
                      new ConstantExpression(new BoolValue(false)),
                      _endIfCode,
                      _endIfTokens)
                      .Evaluate().AsTokens);
            return new TokensValue(tokens);
        }
    }
}
