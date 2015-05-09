﻿using System;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class DeclarationSymbolsListener : VBABaseListener
    {
        private readonly Declarations _declarations = new Declarations();
        public Declarations Declarations { get { return _declarations; } }

        private readonly QualifiedModuleName _qualifiedName;

        private string _currentScope;

        public DeclarationSymbolsListener(VBComponentParseResult result)
            : this(result.QualifiedName, Accessibility.Implicit, result.Component.Type == vbext_ComponentType.vbext_ct_StdModule ? DeclarationType.Module : DeclarationType.Class)
        {
            if (result.Component.Type == vbext_ComponentType.vbext_ct_MSForm)
            {
                DeclareControlsAsMembers(result.Component);
            }
        }

        private DeclarationSymbolsListener(QualifiedModuleName qualifiedName, Accessibility componentAccessibility, DeclarationType declarationType)
        {
            _qualifiedName = qualifiedName;

            SetCurrentScope();
            _declarations.Add(new Declaration(_qualifiedName.QualifyMemberName(_qualifiedName.Component.Name), _qualifiedName.Project.Name, _qualifiedName.Component.Name, false, false, componentAccessibility, declarationType, null, Selection.Home));
        }

        /// <summary>
        /// Scans form designer to create a public, self-assigned field for each control on a form.
        /// </summary>
        /// <remarks>
        /// These declarations are meant to be used to identify control event procedures.
        /// </remarks>
        private void DeclareControlsAsMembers(VBComponent form)
        {
            if (form.Type != vbext_ComponentType.vbext_ct_MSForm)
            {
                throw new InvalidOperationException();
            }

            var designer = form.Designer;
            if (designer == null)
            {
                return;
            }

            foreach (var control in ((dynamic)designer).Controls)
            {
                _declarations.Add(new Declaration(_qualifiedName.QualifyMemberName(control.Name), _currentScope, "Control", true, true, Accessibility.Public, DeclarationType.Control, null, Selection.Home));
            }
        }

        private Declaration CreateDeclaration(string identifierName, string asTypeName, Accessibility accessibility, DeclarationType declarationType, ParserRuleContext context, Selection selection, bool selfAssigned = false, bool withEvents = false)
        {
            return new Declaration(new QualifiedMemberName(_qualifiedName, identifierName), _currentScope, asTypeName, selfAssigned, withEvents, accessibility, declarationType, context, selection);
        }

        /// <summary>
        /// Gets the <c>Accessibility</c> for a procedure member.
        /// </summary>
        private Accessibility GetProcedureAccessibility(VBAParser.VisibilityContext visibilityContext)
        {
            var visibility = visibilityContext == null
                ? "Implicit" // "Public"
                : visibilityContext.GetText();

            return (Accessibility)Enum.Parse(typeof(Accessibility), visibility);
        }

        /// <summary>
        /// Gets the <c>Accessibility</c> for a non-procedure member.
        /// </summary>
        private Accessibility GetMemberAccessibility(VBAParser.VisibilityContext visibilityContext)
        {
            var visibility = visibilityContext == null
                ? "Implicit" // "Private"
                : visibilityContext.GetText();

            return (Accessibility)Enum.Parse(typeof(Accessibility), visibility);
        }

        /// <summary>
        /// Sets current scope to module-level.
        /// </summary>
        private void SetCurrentScope()
        {
            _currentScope = _qualifiedName.ToString();
        }

        /// <summary>
        /// Sets current scope to specified module member.
        /// </summary>
        /// <param name="name">The name of the member owning the current scope.</param>
        private void SetCurrentScope(string name)
        {
            _currentScope = _qualifiedName + "." + name;
        }

        public override void EnterOptionBaseStmt(VBAParser.OptionBaseStmtContext context)
        {
            _declarations.Add(CreateDeclaration(context.GetText(), string.Empty, Accessibility.Implicit, DeclarationType.ModuleOption, context, context.GetSelection()));
        }

        public override void EnterOptionCompareStmt(VBAParser.OptionCompareStmtContext context)
        {
            _declarations.Add(CreateDeclaration(context.GetText(), string.Empty, Accessibility.Implicit, DeclarationType.ModuleOption, context, context.GetSelection()));
        }

        public override void EnterOptionExplicitStmt(VBAParser.OptionExplicitStmtContext context)
        {
            _declarations.Add(CreateDeclaration(context.GetText(), string.Empty, Accessibility.Implicit, DeclarationType.ModuleOption, context, context.GetSelection()));
        }

        public override void ExitOptionPrivateModuleStmt(VBAParser.OptionPrivateModuleStmtContext context)
        {
            _declarations.Add(CreateDeclaration(context.GetText(), string.Empty, Accessibility.Implicit, DeclarationType.ModuleOption, context, context.GetSelection()));
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, null, accessibility, DeclarationType.Procedure, context, context.ambiguousIdentifier().GetSelection()));
            SetCurrentScope(name);
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null 
                ? Tokens.Variant 
                : asTypeClause.type().GetText();

            var declaration = CreateDeclaration(name, asTypeName, accessibility, DeclarationType.Function, context, context.ambiguousIdentifier().GetSelection());
            _declarations.Add(declaration);
            SetCurrentScope(name);
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();

            _declarations.Add(CreateDeclaration(name, asTypeName, accessibility, DeclarationType.PropertyGet, context, context.ambiguousIdentifier().GetSelection()));
            SetCurrentScope(name);
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, null, accessibility, DeclarationType.PropertyLet, context, context.ambiguousIdentifier().GetSelection()));
            SetCurrentScope(name);
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, null, accessibility, DeclarationType.PropertySet, context, context.ambiguousIdentifier().GetSelection()));
            SetCurrentScope(name);
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterEventStmt(VBAParser.EventStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, null, accessibility, DeclarationType.Event, context, context.ambiguousIdentifier().GetSelection()));
            SetCurrentScope(name);
        }

        public override void ExitEventStmt(VBAParser.EventStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterDeclareStmt(VBAParser.DeclareStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var nameContext = context.ambiguousIdentifier();
            var name = nameContext.GetText();

            var hasReturnType = context.FUNCTION() != null;

            var asTypeClause = context.asTypeClause();
            var asTypeName = hasReturnType 
                                ? asTypeClause == null
                                    ? Tokens.Variant
                                    : asTypeClause.type().GetText() 
                                : null;

            var selection = nameContext.GetSelection();

            var declarationType = hasReturnType
                ? DeclarationType.LibraryFunction
                : DeclarationType.LibraryProcedure;

            _declarations.Add(CreateDeclaration(name, asTypeName, accessibility, declarationType, context, selection));
            SetCurrentScope(name);
        }

        public override void ExitDeclareStmt(VBAParser.DeclareStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterArgList(VBAParser.ArgListContext context)
        {
            var args = context.arg();
            foreach (var argContext in args)
            {
                var asTypeClause = argContext.asTypeClause();
                var asTypeName = asTypeClause == null
                    ? Tokens.Variant
                    : asTypeClause.type().GetText();

                var identifier = argContext.ambiguousIdentifier();
                _declarations.Add(CreateDeclaration(identifier.GetText(), asTypeName, Accessibility.Implicit, DeclarationType.Parameter, argContext, identifier.GetSelection()));
            }
        }

        public override void EnterLineLabel(VBAParser.LineLabelContext context)
        {
            _declarations.Add(CreateDeclaration(context.ambiguousIdentifier().GetText(), null, Accessibility.Private, DeclarationType.LineLabel, context, context.ambiguousIdentifier().GetSelection(), true));
        }

        public override void EnterVariableSubStmt(VBAParser.VariableSubStmtContext context)
        {
            var parent = (VBAParser.VariableStmtContext)context.Parent.Parent;
            var accessibility = GetMemberAccessibility(parent.visibility());

            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();

            var withEvents = parent.WITHEVENTS() != null;
            var selfAssigned = asTypeClause != null && asTypeClause.NEW() != null;
            
            _declarations.Add(CreateDeclaration(context.ambiguousIdentifier().GetText(), asTypeName, accessibility, DeclarationType.Variable, context, context.ambiguousIdentifier().GetSelection(), selfAssigned, withEvents));
        }

        public override void EnterConstSubStmt(VBAParser.ConstSubStmtContext context)
        {
            var parent = (VBAParser.ConstStmtContext)context.Parent;
            var accessibility = GetMemberAccessibility(parent.visibility());

            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();

            var identifier = context.ambiguousIdentifier();
            var name = identifier.GetText();
            var value = context.valueStmt().GetText();
            var declaration = new ValuedDeclaration(new QualifiedMemberName(_qualifiedName, name), _currentScope, asTypeName, accessibility, DeclarationType.Constant, value, context, identifier.GetSelection());

            _declarations.Add(declaration);
        }

        public override void EnterTypeStmt(VBAParser.TypeStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, null, accessibility, DeclarationType.UserDefinedType, context, context.ambiguousIdentifier().GetSelection()));
            //SetCurrentScope(name);
        }

        public override void ExitTypeStmt(VBAParser.TypeStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterTypeStmt_Element(VBAParser.TypeStmt_ElementContext context)
        {
            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();

            _declarations.Add(CreateDeclaration(context.ambiguousIdentifier().GetText(), asTypeName, Accessibility.Implicit, DeclarationType.UserDefinedTypeMember, context, context.ambiguousIdentifier().GetSelection()));
        }

        public override void EnterEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, null, accessibility, DeclarationType.Enumeration, context, context.ambiguousIdentifier().GetSelection()));
            //SetCurrentScope(name);
        }

        public override void ExitEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterEnumerationStmt_Constant(VBAParser.EnumerationStmt_ConstantContext context)
        {
            _declarations.Add(CreateDeclaration(context.ambiguousIdentifier().GetText(), null, Accessibility.Implicit, DeclarationType.EnumerationMember, context, context.ambiguousIdentifier().GetSelection()));
        }
    }
}
