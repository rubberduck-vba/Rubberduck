using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionVisitorFactory
    {
        IParseTreeVisitor<IUnreachableCaseInspectionValueResults> Create(RubberduckParserState state);
        //IParseTreeVisitor<IDictionary<ParserRuleContext, IUnreachableCaseInspectionValue>> Create(RubberduckParserState state, string evaluationTypeName = "");
    }

    public class UnreachableCaseInspectionVisitorFactory : IUnreachableCaseInspectionVisitorFactory
    {
        public IParseTreeVisitor<IUnreachableCaseInspectionValueResults> Create(RubberduckParserState state)
        {
            return new UnreachableCaseInspectionValueVisitor(state, new UnreachableCaseInspectionValueFactory());
        }
        //public IParseTreeVisitor<IDictionary<ParserRuleContext, IUnreachableCaseInspectionValue>> Create(RubberduckParserState state, string evaluationTypeName = "")
        //{
        //    return new UnreachableCaseInspectionValueVisitor(state, new IUnreachableCaseInspectionValueFactory());
        //}
    }
}
