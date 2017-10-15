using Antlr4.Runtime;
using System.Collections.Generic;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class ConditionalCompilationConstantExpression : Expression
    {
        private readonly IExpression _tokens;
        private readonly IExpression _identifier;
        private readonly IExpression _expression;
        private readonly SymbolTable<string, IValue> _symbolTable;

        public ConditionalCompilationConstantExpression(
            IExpression tokens,
            IExpression identifier, 
            IExpression expression,
            SymbolTable<string, IValue> symbolTable)
        {
            _tokens = tokens;
            _identifier = identifier;
            _expression = expression;
            _symbolTable = symbolTable;
        }

        public override IValue Evaluate()
        {
            // 3.4.1: If <cc-var-lhs> is a <TYPED-NAME> with a <type-suffix>, the <type-suffix> is ignored.
            var identifier = _identifier.Evaluate().AsString;
            var constantValue = _expression.Evaluate();
            _symbolTable.Add(identifier, constantValue);
            return new LivelinessExpression(
                isAlive: new ConstantExpression(new BoolValue(false)),
                tokens: _tokens).Evaluate();
        }
    }
}
