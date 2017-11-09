using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.VBA
{
    public class AttributeParser : IAttributeParser
    {
        private readonly IModuleExporter _exporter;
        private readonly Func<IVBAPreprocessor> _preprocessorFactory;

        public AttributeParser(IModuleExporter exporter, Func<IVBAPreprocessor> preprocessorFactory)
        {
            _exporter = exporter;
            _preprocessorFactory = preprocessorFactory;
        }

        /// <summary>
        /// Exports the specified component to a temporary file, loads, and then parses the exported file.
        /// </summary>
        /// <param name="module"></param>
        /// <param name="cancellationToken"></param>
        public (IParseTree tree, ITokenStream tokenStream, IDictionary<Tuple<string, DeclarationType>, Attributes> attributes) Parse(QualifiedModuleName module, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var path = _exporter.Export(module.Component);
            if (!File.Exists(path))
            {
                // a document component without any code wouldn't be exported (file would be empty anyway).
                return (null, null, new Dictionary<Tuple<string, DeclarationType>, Attributes>());
            }

            string code;
            if (module.ComponentType == ComponentType.Document)
            {
                code = File.ReadAllText(path, Encoding.UTF8);   //We export the code from Documents as UTF8.
            }
            else
            {
                code = File.ReadAllText(path, Encoding.Default);    //The VBE exports encoded in the current ANSI codepage from the windows settings.
            }

            try
            {
                File.Delete(path);
            }
            catch
            {
                // Meh.
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
