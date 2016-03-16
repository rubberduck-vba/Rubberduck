using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
    /// <summary>
    /// A dictionary storing values for a given attribute.
    /// </summary>
    /// <remarks>
    /// Dictionary key is the attribute name/identifier.
    /// </remarks>
    public class Attributes : Dictionary<string, IEnumerable<string>> { }

    public class AttributeParser : IAttributeParser
    {
        private readonly IModuleExporter _exporter;

        public AttributeParser(IModuleExporter exporter)
        {
            _exporter = exporter;
        }

        /// <summary>
        /// Exports the specified component to a temporary file, loads, and then parses the exported file.
        /// </summary>
        /// <param name="component"></param>
        public IDictionary<Tuple<string, DeclarationType>, Attributes> Parse(VBComponent component)
        {
            var path = _exporter.Export(component);
            if (!File.Exists(path))
            {
                // a document component without any code wouldn't be exported (file would be empty anyway).
                return new Dictionary<Tuple<string, DeclarationType>, Attributes>();
            }

            var code = File.ReadAllText(path);
            File.Delete(path);

            var type = component.Type == vbext_ComponentType.vbext_ct_StdModule
                ? DeclarationType.Module
                : DeclarationType.Class;
            var listener = new AttributeListener(Tuple.Create(component.Name, type));

            var stream = new AntlrInputStream(code);
            var lexer = new AttributesLexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new AttributesParser(tokens);
            parser.AddParseListener(listener);

            // parse tree isn't usable for declarations because
            // line numbers are offset due to module header and attributes
            // (these don't show up in the VBE, that's why we're parsing an exported file)
            var tree = parser.startRule();

            return listener.Attributes;
        }

        private class AttributeListener : AttributesBaseListener
        {
            private readonly Dictionary<Tuple<string, DeclarationType>, Attributes> _attributes = 
                new Dictionary<Tuple<string, DeclarationType>, Attributes>();

            public AttributeListener(Tuple<string,DeclarationType> scope)
            {
                _currentScope = scope;
                _currentScopeAttributes = new Attributes();
            }

            public IDictionary<Tuple<string, DeclarationType>, Attributes> Attributes
            {
                get { return _attributes; }
            }

            private Tuple<string,DeclarationType> _currentScope;
            private Attributes _currentScopeAttributes;

            public override void ExitModuleAttributes(AttributesParser.ModuleAttributesContext context)
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
            }

            public override void EnterSubStmt(AttributesParser.SubStmtContext context)
            {
                _currentScope = Tuple.Create(context.ambiguousIdentifier().GetText(), DeclarationType.Procedure);
                _currentScopeAttributes = new Attributes();
            }

            public override void ExitSubStmt(AttributesParser.SubStmtContext context)
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
            }

            public override void EnterFunctionStmt(AttributesParser.FunctionStmtContext context)
            {
                _currentScope = Tuple.Create(context.ambiguousIdentifier().GetText(), DeclarationType.Function);
                _currentScopeAttributes = new Attributes();
            }

            public override void ExitFunctionStmt(AttributesParser.FunctionStmtContext context)
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
            }

            public override void EnterPropertyGetStmt(AttributesParser.PropertyGetStmtContext context)
            {
                _currentScope = Tuple.Create(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyGet);
                _currentScopeAttributes = new Attributes();
            }

            public override void ExitPropertyGetStmt(AttributesParser.PropertyGetStmtContext context)
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
            }

            public override void EnterPropertyLetStmt(AttributesParser.PropertyLetStmtContext context)
            {
                _currentScope = Tuple.Create(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyLet);
                _currentScopeAttributes = new Attributes();
            }

            public override void ExitPropertyLetStmt(AttributesParser.PropertyLetStmtContext context)
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
            }

            public override void EnterPropertySetStmt(AttributesParser.PropertySetStmtContext context)
            {
                _currentScope = Tuple.Create(context.ambiguousIdentifier().GetText(), DeclarationType.PropertySet);
                _currentScopeAttributes = new Attributes();
            }

            public override void ExitPropertySetStmt(AttributesParser.PropertySetStmtContext context)
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
            }

            public override void ExitAttributeStmt(AttributesParser.AttributeStmtContext context)
            {
                var name = context.implicitCallStmt_InStmt().GetText();
                var values = context.literal().Select(e => e.GetText()).ToList();
                _currentScopeAttributes.Add(name, values);
            }

            public override void ExitModuleConfigElement(AttributesParser.ModuleConfigElementContext context)
            {
                var name = context.ambiguousIdentifier().GetText();
                var values = new[] {context.literal().GetText()};
                _currentScopeAttributes.Add(name, values);
            }
        }
    }
}