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

        private readonly (string scopeIdentifier, DeclarationType scopeType) _moduleScope;
        private (string scopeIdentifier, DeclarationType scopeType) _currentScope;
        private Attributes _currentScopeAttributes;

        public AttributeListener((string scopeIdentifier, DeclarationType scopeType) moduleScope)
        {
            _moduleScope = moduleScope;
            _currentScope = moduleScope;
        }

        public IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> Attributes => _attributes;
        public IDictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext> MembersAllowingAttributes => _membersAllowingAttributes;

        public override void EnterStartRule(VBAParser.StartRuleContext context)
        {
            _membersAllowingAttributes[_moduleScope] = context;
        }

        public override void EnterModuleVariableStmt(VBAParser.ModuleVariableStmtContext context)
        {
            var variableDeclarationStatemenList = context.variableStmt().variableListStmt().variableSubStmt();
            foreach (var variableContext in variableDeclarationStatemenList)
            {
                var variableName = Identifier.GetName(variableContext);
                _membersAllowingAttributes[(variableName, DeclarationType.Variable)] = context;
            }
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            _currentScope = (Identifier.GetName(context.subroutineName()), DeclarationType.Procedure);
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

        private void ResetScope()
        {
            _currentScope = _moduleScope;
            _currentScopeAttributes = null;
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _currentScopeAttributes = new Attributes();
            _currentScope = (Identifier.GetName(context.functionName()), DeclarationType.Function);
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
            var attributeName = context.attributeName().GetText();
            var attributeNameParts = attributeName.Split('.');

            //Module attribute
            if (attributeNameParts.Length == 1)
            {
                AddOrUpdateAttribute(_moduleScope, attributeName, context);
                return;
            }

            var scopeName = attributeNameParts[0]; 

            //Might be an attribute for the enclosing procedure, function or poperty.
            if (_currentScopeAttributes != null && scopeName.Equals(_currentScope.scopeIdentifier, StringComparison.OrdinalIgnoreCase))
            {
                AddOrUpdateAttribute(_currentScopeAttributes, attributeName, context);
                return;
            }

            //Member variable attributes
            var moduleVariableScope = (scopeName, DeclarationType.Variable);
            if (_membersAllowingAttributes.TryGetValue(moduleVariableScope, out _))
            {
                AddOrUpdateAttribute(moduleVariableScope, attributeName, context);
            }
        }

        private void AddOrUpdateAttribute((string scopeName, DeclarationType Variable) moduleVariableScope,
            string attributeName, VBAParser.AttributeStmtContext context)
        {
            if (!_attributes.TryGetValue(moduleVariableScope, out var attributes))
            {
                attributes = new Attributes();
                _attributes.Add(moduleVariableScope, attributes);
            }

            AddOrUpdateAttribute(attributes, attributeName, context);
        }

        private static void AddOrUpdateAttribute(Attributes attributes, string attributeName, VBAParser.AttributeStmtContext context)
        {
            var attribute = attributes.SingleOrDefault(a => a.Name.Equals(attributeName, StringComparison.OrdinalIgnoreCase));
            if (attribute != null)
            {
                attribute.AddContext(context);
                return;
            }

            attributes.Add(new AttributeNode(context));
        }
    }
}