using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValueVisitorFactory
    {
        IParseTreeValueVisitor Create(IReadOnlyList<QualifiedContext<VBAParser.EnumerationStmtContext>> allEnums, Func<QualifiedModuleName, ParserRuleContext, (bool success, IdentifierReference idRef)> idRefRetriever);
    }

    public class ParseTreeValueVisitorFactory : IParseTreeValueVisitorFactory
    {
        private readonly IParseTreeValueFactory _valueFactory;

        public ParseTreeValueVisitorFactory(IParseTreeValueFactory valueFactory)
        {
            _valueFactory = valueFactory;
        }

        public IParseTreeValueVisitor Create(IReadOnlyList<QualifiedContext<VBAParser.EnumerationStmtContext>> allEnums, Func<QualifiedModuleName, ParserRuleContext, (bool success, IdentifierReference idRef)> identifierReferenceRetriever)
        {
            return new ParseTreeValueVisitor(_valueFactory, allEnums, identifierReferenceRetriever);
        }
    }
}