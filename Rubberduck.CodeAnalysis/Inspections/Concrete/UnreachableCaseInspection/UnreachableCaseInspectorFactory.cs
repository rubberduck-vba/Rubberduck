using Antlr4.Runtime;
using System;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectorFactory
    {
        IUnreachableCaseInspector Create(Func<string,QualifiedModuleName,ParserRuleContext,string> func = null);
    }

    public class UnreachableCaseInspectorFactory : IUnreachableCaseInspectorFactory
    {
        private readonly IParseTreeValueFactory _valueFactory;

        public UnreachableCaseInspectorFactory(IParseTreeValueFactory valueFactory)
        {
            _valueFactory = valueFactory;
        }

        public IUnreachableCaseInspector Create(Func<string, QualifiedModuleName, ParserRuleContext, string> func = null)
        {
            return new UnreachableCaseInspector(_valueFactory, func);
        }
    }
}
