using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
    public class AttributeListener : VBAParserBaseListener
    {
        private readonly Dictionary<Tuple<string, DeclarationType>, Attributes> _attributes =
            new Dictionary<Tuple<string, DeclarationType>, Attributes>();

        public AttributeListener(Tuple<string, DeclarationType> scope)
        {
            _currentScope = scope;
            _currentScopeAttributes = new Attributes();
        }

        public IDictionary<Tuple<string, DeclarationType>, Attributes> Attributes => _attributes;

        private IAnnotatedContext _currentAnnotatedContext;
        private Tuple<string, DeclarationType> _currentScope;
        private Attributes _currentScopeAttributes;

        public override void ExitAnnotation(VBAParser.AnnotationContext context)
        {
            _currentAnnotatedContext?.Annotate(context);
            context.AnnotatedContext = _currentAnnotatedContext as ParserRuleContext;
        }

        public override void ExitModuleAttributes(VBAParser.ModuleAttributesContext context)
        {
            if(_currentScopeAttributes.Any() && !_attributes.ContainsKey(_currentScope))
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
            }
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            _currentScope = Tuple.Create(Identifier.GetName(context.subroutineName()), DeclarationType.Procedure);
            _currentAnnotatedContext = context;
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            if(_currentScopeAttributes.Any())
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
                context.AddAttributes(_currentScopeAttributes);
            }
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            _currentScope = Tuple.Create(Identifier.GetName(context.functionName()), DeclarationType.Function);
            _currentAnnotatedContext = context;
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            if(_currentScopeAttributes.Any())
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
                context.AddAttributes(_currentScopeAttributes);
            }
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            _currentScope = Tuple.Create(Identifier.GetName(context.functionName()), DeclarationType.PropertyGet);
            _currentAnnotatedContext = context;
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            if(_currentScopeAttributes.Any())
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
                context.AddAttributes(_currentScopeAttributes);
            }
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            _currentScope = Tuple.Create(Identifier.GetName(context.subroutineName()), DeclarationType.PropertyLet);
            _currentAnnotatedContext = context;
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            if(_currentScopeAttributes.Any())
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
                context.AddAttributes(_currentScopeAttributes);
            }
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            _currentScope = Tuple.Create(Identifier.GetName(context.subroutineName()), DeclarationType.PropertySet);
            _currentAnnotatedContext = context;
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            if(_currentScopeAttributes.Any())
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
                context.AddAttributes(_currentScopeAttributes);
            }
        }

        public override void ExitAttributeStmt(VBAParser.AttributeStmtContext context)
        {
            // We assume attributes can either be simple names (VB_Name) or, if they are inside procedures, member access expressions
            // (e.g. Foo.VB_UserMemId = 0)
            var expr = context.attributeName().lExpression();
            var exprContext = expr as VBAParser.SimpleNameExprContext;
            var name = exprContext != null
                ? exprContext.identifier().GetText()
                : GetAttributeNameWithoutProcedureName((VBAParser.MemberAccessExprContext)expr);

            var values = context.attributeValue().Select(e => e.GetText().Replace("\"", string.Empty)).ToList();
            IEnumerable<string> existingValues;
            if(_currentScopeAttributes.TryGetValue(name, out existingValues))
            {
                values.InsertRange(0, existingValues);
            }
            _currentScopeAttributes[name] = values;
        }

        private static string GetAttributeNameWithoutProcedureName(VBAParser.MemberAccessExprContext expr)
        {
            var name = expr.unrestrictedIdentifier().GetText();
            // The simple name expression represents the procedure's name.
            // We don't want that one though so we simply ignore it.
            if(expr.lExpression() is VBAParser.SimpleNameExprContext)
            {
                return name;
            }
            return $"{GetAttributeNameWithoutProcedureName((VBAParser.MemberAccessExprContext)expr.lExpression())}.{name}";
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