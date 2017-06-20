using System;
using System.Collections.Generic;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    public class TestAttributeParser : IAttributeParser
    {
        private readonly Func<IVBAPreprocessor> _preprocessorFactory;
        public TestAttributeParser(Func<IVBAPreprocessor> preprocessorFactory)
        {
            _preprocessorFactory = preprocessorFactory;
        }

        public IDictionary<Tuple<string, DeclarationType>, Attributes> Parse(IVBComponent component, CancellationToken token, out ITokenStream stream, out IParseTree tree)
        {
            var code = component.CodeModule.Content();
            var type = component.Type == ComponentType.StandardModule
                ? DeclarationType.ProceduralModule
                : DeclarationType.ClassModule;
            var tokenStreamProvider = new SimpleVBAModuleTokenStreamProvider();
            var tokens = tokenStreamProvider.Tokens(code);
            var preprocessor = _preprocessorFactory();
            preprocessor.PreprocessTokenStream(component.Name, tokens, token);
            var listener = new AttributeListener(Tuple.Create(component.Name, type));
            // parse tree isn't usable for declarations because
            // line numbers are offset due to module header and attributes
            // (these don't show up in the VBE, that's why we're parsing an exported file)

            tree = new VBAModuleParser().Parse(component.Name, tokens, new IParseTreeListener[] { listener }, new ExceptionErrorListener(), out stream);

            token.ThrowIfCancellationRequested();
            return listener.Attributes;
        }
    }
}