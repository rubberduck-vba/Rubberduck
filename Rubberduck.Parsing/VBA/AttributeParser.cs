using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using System;
using System.Threading;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.Parsing.VBA
{
    public class AttributeParser : IAttributeParser
    {
        private readonly ISourceCodeProvider _attributesSourceCodeProvider;
        private readonly Func<IVBAPreprocessor> _preprocessorFactory;

        public AttributeParser(ISourceCodeProvider attributesSourceCodeProvider, Func<IVBAPreprocessor> preprocessorFactory)
        {
            _attributesSourceCodeProvider = attributesSourceCodeProvider;
            _preprocessorFactory = preprocessorFactory;
        }

        /// <summary>
        /// Exports the specified component to a temporary file, loads, and then parses the exported file.
        /// </summary>
        /// <param name="module"></param>
        /// <param name="cancellationToken"></param>
        public (IParseTree tree, ITokenStream tokenStream) Parse(QualifiedModuleName module, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var code = _attributesSourceCodeProvider.SourceCode(module);
            cancellationToken.ThrowIfCancellationRequested();

            var tokenStreamProvider = new SimpleVBAModuleTokenStreamProvider();
            var tokens = tokenStreamProvider.Tokens(code);
            var preprocessor = _preprocessorFactory();
            var preprocessorErrorListener = new PreprocessorExceptionErrorListener(module.ComponentName, ParsePass.AttributesPass);
            preprocessor.PreprocessTokenStream(module.ProjectId, module.ComponentName, tokens, preprocessorErrorListener, cancellationToken);
            // parse tree isn't usable for declarations because
            // line numbers are offset due to module header and attributes
            // (these don't show up in the VBE, that's why we're parsing an exported file)

            var mainParseErrorListener = new MainParseExceptionErrorListener(module.ComponentName, ParsePass.AttributesPass);
            var parseResults = new VBAModuleParser().Parse(module.ComponentName, tokens, mainParseErrorListener);

            cancellationToken.ThrowIfCancellationRequested();
            return (parseResults.tree, parseResults.tokenStream);
        }
    }
}
