using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.VBA
{
    public class AttributeParser : IAttributeParser
    {
        private readonly ISourceCodeHandler _sourceCodeHandler;
        private readonly Func<IVBAPreprocessor> _preprocessorFactory;
        private readonly IProjectsProvider _projectsProvider;

        public AttributeParser(ISourceCodeHandler sourceCodeHandler, Func<IVBAPreprocessor> preprocessorFactory, IProjectsProvider projectsProvider)
        {
            _sourceCodeHandler = sourceCodeHandler;
            _preprocessorFactory = preprocessorFactory;
            _projectsProvider = projectsProvider;
        }

        /// <summary>
        /// Exports the specified component to a temporary file, loads, and then parses the exported file.
        /// </summary>
        /// <param name="module"></param>
        /// <param name="cancellationToken"></param>
        public (IParseTree tree, ITokenStream tokenStream, IDictionary<Tuple<string, DeclarationType>, Attributes> attributes) Parse(QualifiedModuleName module, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var component = _projectsProvider.Component(module);

            var code = _sourceCodeHandler.Read(component);
            if (code == null)
            {
                return (null, null, new Dictionary<Tuple<string, DeclarationType>, Attributes>());
            }

            cancellationToken.ThrowIfCancellationRequested();

            var type = module.ComponentType == ComponentType.StandardModule
                ? DeclarationType.ProceduralModule
                : DeclarationType.ClassModule;
            var tokenStreamProvider = new SimpleVBAModuleTokenStreamProvider();
            var tokens = tokenStreamProvider.Tokens(code);
            var preprocessor = _preprocessorFactory();
            var preprocessorErrorListener = new PreprocessorExceptionErrorListener(module.ComponentName, ParsePass.AttributesPass); 
            preprocessor.PreprocessTokenStream(module.ComponentName, tokens, preprocessorErrorListener, cancellationToken);
            var listener = new AttributeListener(Tuple.Create(module.ComponentName, type));
            // parse tree isn't usable for declarations because
            // line numbers are offset due to module header and attributes
            // (these don't show up in the VBE, that's why we're parsing an exported file)

            var mainParseErrorListener = new MainParseExceptionErrorListener(module.ComponentName, ParsePass.AttributesPass);
            var parseResults = new VBAModuleParser().Parse(module.ComponentName, tokens, new IParseTreeListener[] { listener }, mainParseErrorListener);

            cancellationToken.ThrowIfCancellationRequested();
            return (parseResults.tree, parseResults.tokenStream, listener.Attributes);
        }
    }
}
