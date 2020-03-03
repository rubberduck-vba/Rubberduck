using System;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValueVisitorFactory
    {
        IParseTreeValueVisitor Create(Func<QualifiedModuleName, ParserRuleContext, (bool success, IdentifierReference idRef)> idRefRetriever);
    }

    public class ParseTreeValueVisitorFactory : IParseTreeValueVisitorFactory
    {
        private readonly IParseTreeValueFactory _valueFactory;

        public ParseTreeValueVisitorFactory(IParseTreeValueFactory valueFactory)
        {
            _valueFactory = valueFactory;
        }

        public IParseTreeValueVisitor Create(Func<QualifiedModuleName, ParserRuleContext, (bool success, IdentifierReference idRef)> identifierReferenceRetriever)
        {
            return new ParseTreeValueVisitor(_valueFactory, identifierReferenceRetriever);
        }
    }
}