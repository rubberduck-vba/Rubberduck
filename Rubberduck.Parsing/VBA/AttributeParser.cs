using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
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
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);

            // parse tree isn't usable for declarations because
            // line numbers are offset due to module header and attributes
            // (these don't show up in the VBE, that's why we're parsing an exported file)
            var tree = parser.startRule();
            ParseTreeWalker.Default.Walk(listener, tree);

            return listener.Attributes;
        }

        private class AttributeListener : VBAParserBaseListener
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

            public override void ExitModuleAttributes(VBAParser.ModuleAttributesContext context)
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
            }

            private static readonly IReadOnlyDictionary<Type, DeclarationType> ScopingContextTypes =
                new Dictionary<Type, DeclarationType>
            {
                {typeof(VBAParser.SubStmtContext), DeclarationType.Procedure},
                {typeof(VBAParser.FunctionStmtContext), DeclarationType.Function},
                {typeof(VBAParser.PropertyGetStmtContext), DeclarationType.PropertyGet},
                {typeof(VBAParser.PropertyLetStmtContext), DeclarationType.PropertyLet},
                {typeof(VBAParser.PropertySetStmtContext), DeclarationType.PropertySet}
            };

            public override void EnterAmbiguousIdentifier(VBAParser.AmbiguousIdentifierContext context)
            {
                DeclarationType type;
                if (!ScopingContextTypes.TryGetValue(context.Parent.GetType(), out type))
                {
                    return;
                }

                _currentScope = Tuple.Create(context.GetText(), type);
            }

            public override void EnterSubStmt(VBAParser.SubStmtContext context)
            {
                _currentScopeAttributes = new Attributes();
                _currentScope = Tuple.Create(context.ambiguousIdentifier().GetText(), DeclarationType.Procedure);
            }

            public override void ExitSubStmt(VBAParser.SubStmtContext context)
            {
                if (!string.IsNullOrEmpty(_currentScope.Item1) && _currentScopeAttributes.Any())
                {
                    _attributes.Add(_currentScope, _currentScopeAttributes);
                }
            }

            public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
            {
                _currentScopeAttributes = new Attributes();
                _currentScope = Tuple.Create(context.ambiguousIdentifier().GetText(), DeclarationType.Function);
            }

            public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
            {
                if (!string.IsNullOrEmpty(_currentScope.Item1) && _currentScopeAttributes.Any())
                {
                    _attributes.Add(_currentScope, _currentScopeAttributes);
                }
            }

            public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
            {
                _currentScopeAttributes = new Attributes();
                _currentScope = Tuple.Create(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyGet);
            }

            public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
            {
                if (!string.IsNullOrEmpty(_currentScope.Item1) && _currentScopeAttributes.Any())
                {
                    _attributes.Add(_currentScope, _currentScopeAttributes);
                }
            }

            public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
            {
                _currentScopeAttributes = new Attributes();
                _currentScope = Tuple.Create(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyLet);
            }

            public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
            {
                if (!string.IsNullOrEmpty(_currentScope.Item1) && _currentScopeAttributes.Any())
                {
                    _attributes.Add(_currentScope, _currentScopeAttributes);
                }
            }

            public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
            {
                _currentScopeAttributes = new Attributes();
                _currentScope = Tuple.Create(context.ambiguousIdentifier().GetText(), DeclarationType.PropertySet);
            }

            public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
            {
                if (!string.IsNullOrEmpty(_currentScope.Item1) && _currentScopeAttributes.Any())
                {
                    _attributes.Add(_currentScope, _currentScopeAttributes);
                }
            }

            public override void ExitAttributeStmt(VBAParser.AttributeStmtContext context)
            {
                var name = context.implicitCallStmt_InStmt().GetText().Trim();
                var values = context.literal().Select(e => e.GetText().Replace("\"", string.Empty)).ToList();
                _currentScopeAttributes.Add(name, values);
            }

            public override void ExitModuleConfigElement(VBAParser.ModuleConfigElementContext context)
            {
                var name = context.ambiguousIdentifier().GetText();
                var literal = context.literal();
                var values = new[] { literal == null ? string.Empty : literal.GetText()};
                _currentScopeAttributes.Add(name, values);
            }
        }
    }
}