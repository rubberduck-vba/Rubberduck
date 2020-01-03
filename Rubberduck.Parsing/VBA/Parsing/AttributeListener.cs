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

        public override void EnterDeclareStmt(VBAParser.DeclareStmtContext context)
        {
            var name = Identifier.GetName(context.identifier());
            var declarationType = context.FUNCTION() != null
                ? DeclarationType.LibraryFunction
                : DeclarationType.LibraryProcedure;
            var attributeScope = (name, declarationType);
            PushNewScope(attributeScope);
            _membersAllowingAttributes[attributeScope] = context;
        }

        public override void ExitDeclareStmt(VBAParser.DeclareStmtContext context)
        {
            SaveCurrentScopeAttributes(context);
            PopScope();
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            var attributeScope = (Identifier.GetName(context.subroutineName()), DeclarationType.Procedure);
            PushNewScope(attributeScope);
            _membersAllowingAttributes[attributeScope] = context;
        }

        private void PushNewScope((string scopeIdentifier, DeclarationType scopeType) attributeScope)
        {
            _currentScope = attributeScope;
            _currentScopeAttributes = new Attributes();
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            SaveCurrentScopeAttributes(context);
            PopScope();
        }

        private void SaveCurrentScopeAttributes(IAnnotatedContext context)
        {
            if (_currentScopeAttributes.Any())
            {
                _attributes.Add(_currentScope, _currentScopeAttributes);
                context.AddAttributes(_currentScopeAttributes);
            }
        }

        private void PopScope()
        {
            _currentScope = _moduleScope;
            _currentScopeAttributes = null;
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            var attributeScope = (Identifier.GetName(context.functionName()), DeclarationType.Function);
            PushNewScope(attributeScope);
            _membersAllowingAttributes[attributeScope] = context;
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SaveCurrentScopeAttributes(context);
            PopScope();
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            var attributeScope = (Identifier.GetName(context.functionName()), DeclarationType.PropertyGet);
            PushNewScope(attributeScope);
            _membersAllowingAttributes[attributeScope] = context;
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SaveCurrentScopeAttributes(context);
            PopScope();
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            var attributeScope = (Identifier.GetName(context.subroutineName()), DeclarationType.PropertyLet);
            PushNewScope(attributeScope);
            _membersAllowingAttributes[attributeScope] = context;
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SaveCurrentScopeAttributes(context);
            PopScope();
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            var attributeScope = (Identifier.GetName(context.subroutineName()), DeclarationType.PropertySet);
            PushNewScope(attributeScope);
            _membersAllowingAttributes[attributeScope] = context;
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SaveCurrentScopeAttributes(context);
            PopScope();
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
            attributes.Add(new AttributeNode(context));
        }
    }
}