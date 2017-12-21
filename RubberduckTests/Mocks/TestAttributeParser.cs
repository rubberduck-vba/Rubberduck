using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Mocks
{
    public class TestAttributeParser : IAttributeParser
    {
        private readonly Func<IVBAPreprocessor> _preprocessorFactory;
        private readonly RubberduckParserState _state;

        public TestAttributeParser(Func<IVBAPreprocessor> preprocessorFactory, RubberduckParserState state)
        {
            _preprocessorFactory = preprocessorFactory;
            _state = state;
        }

        public (IParseTree tree, ITokenStream tokenStream, IDictionary<Tuple<string, DeclarationType>, Attributes> attributes) Parse(QualifiedModuleName module, CancellationToken cancellationToken)
        {
            var code = module.Component.CodeModule.Content();
            var type = module.ComponentType == ComponentType.StandardModule
                ? DeclarationType.ProceduralModule
                : DeclarationType.ClassModule;
            var tokenStreamProvider = new SimpleVBAModuleTokenStreamProvider();
            var tokens = tokenStreamProvider.Tokens(code);
            var preprocessor = _preprocessorFactory();
            var preprocessingErrorListener = new PreprocessorExceptionErrorListener(module.ComponentName, ParsePass.AttributesPass);
            preprocessor.PreprocessTokenStream(module.ComponentName, tokens, preprocessingErrorListener, cancellationToken);
            var listener = new AttributeListener(Tuple.Create(module.ComponentName, type));
            // parse tree isn't usable for declarations because
            // line numbers are offset due to module header and attributes
            // (these don't show up in the VBE, that's why we're parsing an exported file)

            var mainParseErrorListener = new SyntaxErrorNotificationListener(module);
            mainParseErrorListener.OnSyntaxError += (sender, e) =>
            {
                _state.AddParserError(e);
            };

            var parseResults = new VBAModuleParser().Parse(module.Name, PredictionMode.Sll, tokens, new IParseTreeListener[] { listener }, mainParseErrorListener);
            if (_state.ModuleExceptions.Any(r => r.Item1 == module))
            {
                _state.ClearExceptions(module);
                parseResults = new VBAModuleParser().Parse(module.Name, PredictionMode.Ll, tokens, new IParseTreeListener[] {listener},
                    mainParseErrorListener);
            }

            cancellationToken.ThrowIfCancellationRequested();
            return (parseResults.tree, parseResults.tokenStream, listener.Attributes);
        }
    }
}