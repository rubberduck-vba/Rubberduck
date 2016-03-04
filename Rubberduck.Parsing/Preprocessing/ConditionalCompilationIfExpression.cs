using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class ConditionalCompilationIfExpression : Expression
    {
        private readonly IExpression _ifCondCode;
        private readonly IExpression _ifCond;
        private readonly IExpression _ifBlock;
        private readonly IEnumerable<Tuple<IExpression, IExpression, IExpression>> _elseIfCodeCondBlocks;
        private readonly IExpression _elseCondCode;
        private readonly IExpression _elseBlock;
        private readonly IExpression _endIfCode;

        public ConditionalCompilationIfExpression(
            IExpression ifCondCode,
            IExpression ifCond,
            IExpression ifBlock,
            IEnumerable<Tuple<IExpression, IExpression, IExpression>> elseIfCodeCondBlocks,
            IExpression elseCondCode,
            IExpression elseBlock,
            IExpression endIfCode)
        {
            _ifCondCode = ifCondCode;
            _ifCond = ifCond;
            _ifBlock = ifBlock;
            _elseIfCodeCondBlocks = elseIfCodeCondBlocks;
            _elseCondCode = elseCondCode;
            _elseBlock = elseBlock;
            _endIfCode = endIfCode;
        }

        public override IValue Evaluate()
        {
            StringBuilder builder = new StringBuilder();
            List<bool> conditions = new List<bool>();
            builder.Append(
                new LivelinessExpression(
                    new ConstantExpression(new BoolValue(false)),
                    _ifCondCode)
                    .Evaluate().AsString);

            var ifIsAlive = _ifCond.EvaluateCondition();
            conditions.Add(ifIsAlive);
            builder.Append(
                new LivelinessExpression(
                    new ConstantExpression(new BoolValue(ifIsAlive)),
                    _ifBlock)
                    .Evaluate().AsString);

            foreach (var elseIf in _elseIfCodeCondBlocks)
            {
                builder.Append(
                   new LivelinessExpression(
                       new ConstantExpression(new BoolValue(false)),
                       elseIf.Item1)
                       .Evaluate().AsString);
                var elseIfIsAlive = !ifIsAlive && elseIf.Item2.EvaluateCondition();
                conditions.Add(elseIfIsAlive);
                builder.Append(
                    new LivelinessExpression(
                        new ConstantExpression(new BoolValue(elseIfIsAlive)),
                        elseIf.Item3)
                        .Evaluate().AsString);
            }

            if (_elseCondCode != null)
            {
                builder.Append(
                   new LivelinessExpression(
                       new ConstantExpression(new BoolValue(false)),
                       _elseCondCode)
                       .Evaluate().AsString);
                var elseIsAlive = conditions.All(condition => !condition);
                builder.Append(
                    new LivelinessExpression(
                        new ConstantExpression(new BoolValue(elseIsAlive)),
                        _elseBlock)
                        .Evaluate().AsString);
            }
            builder.Append(
                  new LivelinessExpression(
                      new ConstantExpression(new BoolValue(false)),
                      _endIfCode)
                      .Evaluate().AsString);
            return new StringValue(builder.ToString());
        }
    }
}
