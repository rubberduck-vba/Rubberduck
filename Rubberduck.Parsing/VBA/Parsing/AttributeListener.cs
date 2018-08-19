using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public class AttributeListener : VBAParserBaseListener
    {
        private readonly Dictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> _attributes 
            = new Dictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes>();
        private readonly Dictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext> _membersAllowingAttributes
            = new Dictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext>();

        private readonly (string scopeIdentifier, DeclarationType scopeType) _initialScope;
        private (string scopeIdentifier, DeclarationType scopeType) _currentScope;
        private IAnnotatedContext _currentAnnotatedContext;
        private Attributes _currentScopeAttributes;

        public AttributeListener((string scopeIdentifier, DeclarationType scopeType) initialScope)
        {
            _initialScope = initialScope;
            _currentScope = initialScope;
            _currentScopeAttributes = new Attributes();
        }

        public IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> Attributes => _attributes;
        public IDictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext> MembersAllowingAttributes => _membersAllowingAttributes;

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

        public override void EnterModuleVariableStmt(VBAParser.ModuleVariableStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            var annotatedContext = context.variableStmt().variableListStmt().variableSubStmt().Last();
            _currentScope = (Identifier.GetName(annotatedContext), DeclarationType.Variable);
            _currentAnnotatedContext = annotatedContext;
            _membersAllowingAttributes[_currentScope] = context;
        }

        public override void ExitModuleVariableStmt(VBAParser.ModuleVariableStmtContext context)
        {
            if (_currentScopeAttributes.Any())
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
                var annotatedContext = context.variableStmt().variableListStmt().variableSubStmt().Last();
                annotatedContext.AddAttributes(_currentScopeAttributes);
            }

            ResetScope();
        }

        private void ResetScope()
        {
            _currentScope = _initialScope;
            _currentScopeAttributes = _attributes.TryGetValue(_currentScope, out var attributes)
                ? attributes
                : new Attributes();
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            _currentScope = (Identifier.GetName(context.subroutineName()), DeclarationType.Procedure);
            _currentAnnotatedContext = context;
            _membersAllowingAttributes[_currentScope] = context;
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            if(_currentScopeAttributes.Any())
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
                context.AddAttributes(_currentScopeAttributes);
            }

            ResetScope();
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            _currentScope = (Identifier.GetName(context.functionName()), DeclarationType.Function);
            _currentAnnotatedContext = context;
            _membersAllowingAttributes[_currentScope] = context;
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            if(_currentScopeAttributes.Any())
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
                context.AddAttributes(_currentScopeAttributes);
            }

            ResetScope();
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            _currentScope = (Identifier.GetName(context.functionName()), DeclarationType.PropertyGet);
            _currentAnnotatedContext = context;
            _membersAllowingAttributes[_currentScope] = context;
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            if(_currentScopeAttributes.Any())
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
                context.AddAttributes(_currentScopeAttributes);
            }

            ResetScope();
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            _currentScope = (Identifier.GetName(context.subroutineName()), DeclarationType.PropertyLet);
            _currentAnnotatedContext = context;
            _membersAllowingAttributes[_currentScope] = context;
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            if(_currentScopeAttributes.Any())
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
                context.AddAttributes(_currentScopeAttributes);
            }

            ResetScope();
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            _currentScope = (Identifier.GetName(context.subroutineName()), DeclarationType.PropertySet);
            _currentAnnotatedContext = context;
            _membersAllowingAttributes[_currentScope] = context;
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            if(_currentScopeAttributes.Any())
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
                context.AddAttributes(_currentScopeAttributes);
            }

            ResetScope();
        }

        public override void ExitAttributeStmt(VBAParser.AttributeStmtContext context)
        {
            var values = context.attributeValue().Select(e => e.GetText().Replace("\"", string.Empty)).ToList();

            var attribute = _currentScopeAttributes
                .SingleOrDefault(a => a.Name.Equals(context.attributeName().GetText(), StringComparison.OrdinalIgnoreCase));
            if (attribute != null)
            {
                foreach(var value in values.Where(v => !attribute.HasValue(v)))
                {
                    attribute.AddValue(value);
                }
            }

            _currentScopeAttributes.Add(new AttributeNode(context));
        }
    }
}