using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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
        /// <param name="component"></param>
        /// <param name="token"></param>
        /// <param name="stream"></param>
        public IDictionary<Tuple<string, DeclarationType>, Attributes> Parse(IVBComponent component, CancellationToken token, out ITokenStream stream, out IParseTree tree)
        {
            token.ThrowIfCancellationRequested();
            var path = _exporter.Export(component);
            if (!File.Exists(path))
            {
                // a document component without any code wouldn't be exported (file would be empty anyway).
                stream = null;
                tree = null;
                return new Dictionary<Tuple<string, DeclarationType>, Attributes>();
            }
            var code = File.ReadAllText(path);
            try
            {
                File.Delete(path);
            }
            catch
            {
                // Meh.
            }
           
            token.ThrowIfCancellationRequested();

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
