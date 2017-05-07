using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        public IDictionary<Tuple<string, DeclarationType>, Attributes> Parse(IVBComponent component, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var path = _exporter.Export(component, true);
            if (!File.Exists(path))
            {
                // a document component without any code wouldn't be exported (file would be empty anyway).
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

            ITokenStream tokenStream;
            new VBAModuleParser().Parse(component.Name, tokens, new IParseTreeListener[] { listener }, new ExceptionErrorListener(), out tokenStream);
            return listener.Attributes;
        }

        private class AttributeListener : VBAParserBaseListener
        {
            private readonly Dictionary<Tuple<string, DeclarationType>, Attributes> _attributes =
                new Dictionary<Tuple<string, DeclarationType>, Attributes>();

            public AttributeListener(Tuple<string, DeclarationType> scope)
            {
                _currentScope = scope;
                _currentScopeAttributes = new Attributes();
            }

            public IDictionary<Tuple<string, DeclarationType>, Attributes> Attributes
            {
                get { return _attributes; }
            }

            private Tuple<string, DeclarationType> _currentScope;
            private Attributes _currentScopeAttributes;

            public override void ExitModuleAttributes(VBAParser.ModuleAttributesContext context)
            {
                if (_currentScopeAttributes.Any() && !_attributes.ContainsKey(_currentScope))
                {
                    _attributes.Add(_currentScope, _currentScopeAttributes);
                }
            }

            public override void EnterSubStmt(VBAParser.SubStmtContext context)
            {
                _currentScopeAttributes = new Attributes();
                _currentScope = Tuple.Create(Identifier.GetName(context.subroutineName()), DeclarationType.Procedure);
            }

            public override void ExitSubStmt(VBAParser.SubStmtContext context)
            {
                if (_currentScopeAttributes.Any())
                {
                    _attributes.Add(_currentScope, _currentScopeAttributes);
                }
            }

            public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
            {
                _currentScopeAttributes = new Attributes();
                _currentScope = Tuple.Create(Identifier.GetName(context.functionName()), DeclarationType.Function);
            }

            public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
            {
                if (_currentScopeAttributes.Any())
                {
                    _attributes.Add(_currentScope, _currentScopeAttributes);
                }
            }

            public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
            {
                _currentScopeAttributes = new Attributes();
                _currentScope = Tuple.Create(Identifier.GetName(context.functionName()), DeclarationType.PropertyGet);
            }

            public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
            {
                if (_currentScopeAttributes.Any())
                {
                    _attributes.Add(_currentScope, _currentScopeAttributes);
                }
            }

            public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
            {
                _currentScopeAttributes = new Attributes();
                _currentScope = Tuple.Create(Identifier.GetName(context.subroutineName()), DeclarationType.PropertyLet);
            }

            public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
            {
                if (_currentScopeAttributes.Any())
                {
                    _attributes.Add(_currentScope, _currentScopeAttributes);
                }
            }

            public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
            {
                _currentScopeAttributes = new Attributes();
                _currentScope = Tuple.Create(Identifier.GetName(context.subroutineName()), DeclarationType.PropertySet);
            }

            public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
            {
                if (_currentScopeAttributes.Any())
                {
                    _attributes.Add(_currentScope, _currentScopeAttributes);
                }
            }

            public override void ExitAttributeStmt(VBAParser.AttributeStmtContext context)
            {
                // We assume attributes can either be simple names (VB_Name) or, if they are inside procedures, member access expressions
                // (e.g. Foo.VB_UserMemId = 0)
                var expr = context.attributeName().lExpression();
                string name;
                if (expr is VBAParser.SimpleNameExprContext)
                {
                    name = ((VBAParser.SimpleNameExprContext)expr).identifier().GetText();
                }
                else
                {
                    // Turns "Foo.VB_ProcData.VB_Invoke_Func" into "VB_ProcData.VB_Invoke_Func",
                    // because we are not interested in the subroutine name Foo.
                    name = GetAttributeNameWithoutProcedureName((VBAParser.MemberAccessExprContext)expr);
                }
                var values = context.attributeValue().Select(e => e.GetText().Replace("\"", string.Empty)).ToList();
                IEnumerable<string> existingValues;
                if (_currentScopeAttributes.TryGetValue(name, out existingValues))
                {
                    values.InsertRange(0, existingValues);
                }
                _currentScopeAttributes[name] = values;
            }

            private string GetAttributeNameWithoutProcedureName(VBAParser.MemberAccessExprContext expr)
            {
                string name = expr.unrestrictedIdentifier().GetText();
                // The simple name expression represents the procedure's name.
                // We don't want that one though so we simply ignore it.
                if (expr.lExpression() is VBAParser.SimpleNameExprContext)
                {
                    return name;
                }
                return string.Format("{0}.{1}", GetAttributeNameWithoutProcedureName((VBAParser.MemberAccessExprContext)expr.lExpression()), name);
            }

            public override void ExitModuleConfigElement(VBAParser.ModuleConfigElementContext context)
            {
                var name = context.unrestrictedIdentifier().GetText();
                var literal = context.expression();
                var values = new[] { literal == null ? string.Empty : literal.GetText() };
                _currentScopeAttributes.Add(name, values);
            }
        }
    }
}
