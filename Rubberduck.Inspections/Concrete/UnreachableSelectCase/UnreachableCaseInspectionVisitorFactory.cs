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
        IParseTreeVisitor<IUnreachableCaseInspectionValue> Create(RubberduckParserState state, string evaluationTypeName = "");
    }

    public class UnreachableCaseInspectionVisitorFactory : IUnreachableCaseInspectionVisitorFactory
    {
        public IParseTreeVisitor<IUnreachableCaseInspectionValue> Create(RubberduckParserState state, string evaluationTypeName = "")
        {
            return new UnreachableCaseInspectionValueVisitor(state, new IUnreachableCaseInspectionValueFactory(), evaluationTypeName ?? string.Empty);
        }
    }
}
