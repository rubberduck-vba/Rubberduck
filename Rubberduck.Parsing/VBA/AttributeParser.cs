using System.Collections.Generic;
using System.IO;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.VBA
{
    public class AttributeParser : IAttributeParser
    {
        private readonly IModuleExporter _exporter;

        public AttributeParser(IModuleExporter exporter)
        {
            _exporter = exporter;
        }
        
        public IDictionary<string, IEnumerable<string>> Parse(VBComponent component)
        {
            var path = _exporter.Export(component);
            var code = File.ReadAllText(path);
            File.Delete(path);

            var listener = new AttributeListener();

            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            parser.AddParseListener(listener);

            // parse tree isn't usable for declarations because
            // line numbers are offset due to module header and attributes
            // (these don't show up in the VBE, that's why we're parsing an exported file)
            var tree = parser.startRule();

            return listener.Attributes;
        }

        private class AttributeListener : VBABaseListener
        {
            private readonly Dictionary<string,IEnumerable<string>> _attributes = new Dictionary<string,IEnumerable<string>>(); 
            public IDictionary<string,IEnumerable<string>> Attributes { get { return _attributes; } }

            public override void ExitAttributeStmt(VBAParser.AttributeStmtContext context)
            {
                var name = context.implicitCallStmt_InStmt().GetText();
                var values = context.literal().Select(e => e.GetText()).ToList();
                _attributes.Add(name, values);
            }

            public override void ExitModuleConfigElement(VBAParser.ModuleConfigElementContext context)
            {
                var name = context.ambiguousIdentifier().GetText();
                var value = context.literal().GetText();
                _attributes.Add(name, new[] {value});
            }
        }
    }
}