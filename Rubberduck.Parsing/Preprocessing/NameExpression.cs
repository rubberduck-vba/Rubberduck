namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class NameExpression : Expression
    {
        private readonly IExpression _identifier;
        private readonly SymbolTable<string, IValue> _symbolTable;

        public NameExpression(
            IExpression identifier,
            SymbolTable<string, IValue> symbolTable)
        {
            _identifier = identifier;
            _symbolTable = symbolTable;
        }

        public override IValue Evaluate()
        {
            var identifier = _identifier.Evaluate().AsString;
            // Special case, identifier that does not exist is Empty.
            // Could add them to the symbol table, but since they are all constants
            // they never change anyway.
            if (!_symbolTable.HasSymbol(identifier))
            {
                return EmptyValue.Value;
            }
            return _symbolTable.Get(identifier);
        }
    }
}
