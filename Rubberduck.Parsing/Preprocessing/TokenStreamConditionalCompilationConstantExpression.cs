using Antlr4.Runtime;
using System.Collections.Generic;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class TokenStreamConditionalCompilationConstantExpression : Expression
    {
        private readonly IExpression _tokenText;
        private readonly IExpression _identifier;
        private readonly IExpression _expression;
        private readonly IEnumerable<IToken> _tokens;
        private readonly SymbolTable<string, IValue> _symbolTable;

        public TokenStreamConditionalCompilationConstantExpression(
            IExpression tokenText,
            IExpression identifier, 
            IExpression expression,
            IEnumerable<IToken> tokens,
            SymbolTable<string, IValue> symbolTable)
        {
            _tokenText = tokenText;
            _identifier = identifier;
            _expression = expression;
            _tokens = tokens;
            _symbolTable = symbolTable;
        }

        public override IValue Evaluate()
        {
            // 3.4.1: If <cc-var-lhs> is a <TYPED-NAME> with a <type-suffix>, the <type-suffix> is ignored.
            var identifier = _identifier.Evaluate().AsString;
            var constantValue = _expression.Evaluate();
            _symbolTable.Add(identifier, constantValue);
            return new TokenStreamLivelinessExpression(
                isAlive: new ConstantExpression(new BoolValue(false)),
                code: _tokenText,
                tokens: _tokens).Evaluate();
        }
    }
}
