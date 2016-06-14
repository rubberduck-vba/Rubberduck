﻿using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Microsoft.Vbe.Interop.Forms;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime.Misc;

namespace Rubberduck.Parsing.Symbols
{
    public class DeclarationSymbolsListener : VBAParserBaseListener
    {
        private readonly QualifiedModuleName _qualifiedName;
        private readonly Declaration _moduleDeclaration;
        private readonly Declaration _projectDeclaration;

        private string _currentScope;
        private Declaration _currentScopeDeclaration;
        private Declaration _parentDeclaration;

        private readonly IEnumerable<CommentNode> _comments;
        private readonly IEnumerable<IAnnotation> _annotations;
        private readonly IDictionary<Tuple<string, DeclarationType>, Attributes> _attributes;
        private readonly HashSet<ReferencePriorityMap> _projectReferences;

        private readonly List<Declaration> _createdDeclarations = new List<Declaration>();
        public IReadOnlyList<Declaration> CreatedDeclarations { get { return _createdDeclarations; } }

        public DeclarationSymbolsListener(
            QualifiedModuleName qualifiedName,
            vbext_ComponentType type,
            IEnumerable<CommentNode> comments,
            IEnumerable<IAnnotation> annotations,
            IDictionary<Tuple<string, DeclarationType>, Attributes> attributes,
            HashSet<ReferencePriorityMap> projectReferences,
            Declaration projectDeclaration)
        {
            _qualifiedName = qualifiedName;
            _comments = comments;
            _annotations = annotations;
            _attributes = attributes;

            var declarationType = type == vbext_ComponentType.vbext_ct_StdModule
                ? DeclarationType.ProceduralModule
                : DeclarationType.ClassModule;

            _projectReferences = projectReferences;
            _projectDeclaration = projectDeclaration;

            var key = Tuple.Create(_qualifiedName.ComponentName, declarationType);
            var moduleAttributes = attributes.ContainsKey(key)
                ? attributes[key]
                : new Attributes();

            if (declarationType == DeclarationType.ProceduralModule)
            {
                _moduleDeclaration = new ProceduralModuleDeclaration(
                    _qualifiedName.QualifyMemberName(_qualifiedName.Component.Name),
                    _projectDeclaration,
                    _qualifiedName.Component.Name,
                    false,
                    FindAnnotations(),
                    moduleAttributes);
            }
            else
            {
                bool hasDefaultInstanceVariable = type != vbext_ComponentType.vbext_ct_ClassModule && type != vbext_ComponentType.vbext_ct_StdModule;
                _moduleDeclaration = new ClassModuleDeclaration(
                    _qualifiedName.QualifyMemberName(_qualifiedName.Component.Name),
                    _projectDeclaration,
                    _qualifiedName.Component.Name,
                    false,
                    FindAnnotations(),
                    moduleAttributes,
                    hasDefaultInstanceVariable: hasDefaultInstanceVariable);
            }
            SetCurrentScope();
            AddDeclaration(_moduleDeclaration);
            var component = _moduleDeclaration.QualifiedName.QualifiedModuleName.Component;
            if (component.Type == vbext_ComponentType.vbext_ct_MSForm || component.Designer != null)
            {
                DeclareControlsAsMembers(component);
            }
        }

        private IEnumerable<IAnnotation> FindAnnotations()
        {
            if (_annotations == null)
            {
                return null;
            }
            var lastDeclarationsSectionLine = _qualifiedName.Component.CodeModule.CountOfDeclarationLines;
            var annotations = _annotations.Where(annotation => annotation.QualifiedSelection.QualifiedName.Equals(_qualifiedName)
                && annotation.QualifiedSelection.Selection.EndLine <= lastDeclarationsSectionLine);
            return annotations.ToList();
        }

        private IEnumerable<IAnnotation> FindAnnotations(int line)
        {
            if (_annotations == null)
            {
                return null;
            }

            var annotations = new List<IAnnotation>();

            // VBE 1-based indexing
            for (var i = line - 1; i >= 1; i--)
            {
                var annotation = _annotations.SingleOrDefault(a => a.QualifiedSelection.Selection.StartLine == i);

                if (annotation == null)
                {
                    break;
                }

                annotations.Add(annotation);
            }

            return annotations;
        }

        /// <summary>
        /// Scans form designer to create a public, self-assigned field for each control on a form.
        /// </summary>
        /// <remarks>
        /// These declarations are meant to be used to identify control event procedures.
        /// </remarks>
        private void DeclareControlsAsMembers(VBComponent form)
        {
            var designer = form.Designer;
            if (designer == null)
            {
                return;
            }
            if (!(designer is UserForm))
            {
                return;
            }
            // "using dynamic typing here, because not only MSForms could have a Controls collection (e.g. MS-Access forms are 'document' modules)."
            // Note: Dynamic doesn't seem to support explicit interfaces that's why we cast it anyway, MS Access forms apparently have to be treated specially anyway.
            var userForm = (UserForm)designer;
            // Not all objects in the Controls collection are Control, some are Images.
            foreach (dynamic control in userForm.Controls)
            {
                // The as type declaration should be TextBox, CheckBox, etc. depending on the type.
                var declaration = new Declaration(
                    _qualifiedName.QualifyMemberName(control.Name),
                    _parentDeclaration,
                    _currentScopeDeclaration,
                    "Control",
                    null,
                    true,
                    true,
                    Accessibility.Public,
                    DeclarationType.Control,
                    null,
                    Selection.Home,
                    false,
                    null,
                    false);
                AddDeclaration(declaration);
            }
        }

        private Declaration CreateDeclaration(
            string identifierName,
            string asTypeName,
            Accessibility accessibility,
            DeclarationType declarationType,
            ParserRuleContext context,
            Selection selection,
            bool isArray,
            VBAParser.AsTypeClauseContext asTypeContext,
            string typeHint,
            bool selfAssigned = false,
            bool withEvents = false)
        {
            Declaration result;
            if (declarationType == DeclarationType.Parameter)
            {
                var argContext = (VBAParser.ArgContext)context;
                var isOptional = argContext.OPTIONAL() != null;
                // TODO: "As Type" could be missing. Temp solution until default values are parsed correctly.
                if (isOptional && asTypeContext != null)
                {
                    // if parameter is optional, asTypeName may contain the default value
                    var complexType = asTypeContext.type().complexType();
                    if (complexType != null && complexType.expression() is VBAParser.RelationalOpContext)
                    {
                        asTypeName = complexType.expression().GetChild(0).GetText();
                    }
                }

                var isByRef = argContext.BYREF() != null;
                var isParamArray = argContext.PARAMARRAY() != null;
                result = new ParameterDeclaration(
                    new QualifiedMemberName(_qualifiedName, identifierName),
                    _parentDeclaration,
                    context,
                    selection,
                    asTypeName,
                    asTypeContext,
                    typeHint,
                    isOptional,
                    isByRef,
                    isArray,
                    isParamArray);
                if (_parentDeclaration is IDeclarationWithParameter)
                {
                    ((IDeclarationWithParameter)_parentDeclaration).AddParameter(result);
                }
            }
            else
            {
                var key = Tuple.Create(identifierName, declarationType);
                Attributes attributes = null;
                if (_attributes.ContainsKey(key))
                {
                    attributes = _attributes[key];
                }

                var annotations = FindAnnotations(selection.StartLine);
                if (declarationType == DeclarationType.Procedure)
                {
                    result = new SubroutineDeclaration(new QualifiedMemberName(_qualifiedName, identifierName), _parentDeclaration, _currentScopeDeclaration, asTypeName, accessibility, context, selection, false, annotations, attributes);
                }
                else if (declarationType == DeclarationType.Function)
                {
                    result = new FunctionDeclaration(
                        new QualifiedMemberName(_qualifiedName, identifierName),
                        _parentDeclaration,
                        _currentScopeDeclaration,
                        asTypeName,
                        asTypeContext,
                        typeHint,
                        accessibility,
                        context,
                        selection,
                        isArray,
                        false,
                        annotations,
                        attributes);
                }
                else if (declarationType == DeclarationType.LibraryProcedure || declarationType == DeclarationType.LibraryFunction)
                {
                    result = new ExternalProcedureDeclaration(new QualifiedMemberName(_qualifiedName, identifierName), _parentDeclaration, _currentScopeDeclaration, declarationType, asTypeName, asTypeContext, accessibility, context, selection, false, annotations);
                }
                else if (declarationType == DeclarationType.PropertyGet)
                {
                    result = new PropertyGetDeclaration(
                        new QualifiedMemberName(_qualifiedName, identifierName),
                        _parentDeclaration,
                        _currentScopeDeclaration,
                        asTypeName,
                        asTypeContext,
                        typeHint,
                        accessibility,
                        context,
                        selection,
                        isArray,
                        false,
                        annotations,
                        attributes);
                }
                else if (declarationType == DeclarationType.PropertySet)
                {
                    result = new PropertySetDeclaration(new QualifiedMemberName(_qualifiedName, identifierName), _parentDeclaration, _currentScopeDeclaration, asTypeName, accessibility, context, selection, false, annotations, attributes);
                }
                else if (declarationType == DeclarationType.PropertyLet)
                {
                    result = new PropertyLetDeclaration(new QualifiedMemberName(_qualifiedName, identifierName), _parentDeclaration, _currentScopeDeclaration, asTypeName, accessibility, context, selection, false, annotations, attributes);
                }
                else
                {
                    result = new Declaration(
                        new QualifiedMemberName(_qualifiedName, identifierName),
                        _parentDeclaration,
                        _currentScopeDeclaration,
                        asTypeName,
                        typeHint,
                        selfAssigned,
                        withEvents,
                        accessibility,
                        declarationType,
                        context,
                        selection,
                        isArray,
                        asTypeContext,
                        false,
                        annotations,
                        attributes);
                }
                if (_parentDeclaration.DeclarationType == DeclarationType.ClassModule && result is ICanBeDefaultMember && ((ICanBeDefaultMember)result).IsDefaultMember)
                {
                    ((ClassModuleDeclaration)_parentDeclaration).DefaultMember = result;
                }
            }
            return result;
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

            return (Accessibility)Enum.Parse(typeof(Accessibility), visibility, true);
        }

        /// <summary>
        /// Sets current scope to module-level.
        /// </summary>
        private void SetCurrentScope()
        {
            _currentScope = _qualifiedName.ToString();
            _currentScopeDeclaration = _moduleDeclaration;
            _parentDeclaration = _moduleDeclaration;
        }

        /// <summary>
        /// Sets current scope to specified module member.
        /// </summary>
        /// <param name="procedureDeclaration"></param>
        /// <param name="name">The name of the member owning the current scope.</param>
        private void SetCurrentScope(Declaration procedureDeclaration, string name)
        {
            _currentScope = _qualifiedName + "." + name;
            _currentScopeDeclaration = procedureDeclaration;
            _parentDeclaration = procedureDeclaration;
        }

        public override void EnterImplementsStmt(VBAParser.ImplementsStmtContext context)
        {
            // The expression will be later resolved to the actual declaration. Have to split the work up because we have to gather/create all declarations first.
            ((ClassModuleDeclaration)_moduleDeclaration).AddSupertype(context.expression().GetText());
        }

        public override void EnterOptionBaseStmt(VBAParser.OptionBaseStmtContext context)
        {
            AddDeclaration(CreateDeclaration(
                context.GetText(),
                string.Empty,
                Accessibility.Implicit,
                DeclarationType.ModuleOption,
                context,
                context.GetSelection(),
                false,
                null,
                null));
        }

        public override void EnterOptionCompareStmt(VBAParser.OptionCompareStmtContext context)
        {
            AddDeclaration(CreateDeclaration(
                context.GetText(),
                string.Empty,
                Accessibility.Implicit,
                DeclarationType.ModuleOption,
                context,
                context.GetSelection(),
                false,
                null,
                null));
        }

        public override void EnterOptionExplicitStmt(VBAParser.OptionExplicitStmtContext context)
        {
            AddDeclaration(CreateDeclaration(
                context.GetText(),
                string.Empty,
                Accessibility.Implicit,
                DeclarationType.ModuleOption,
                context,
                context.GetSelection(),
                false,
                null,
                null));
        }

        public override void ExitOptionPrivateModuleStmt(VBAParser.OptionPrivateModuleStmtContext context)
        {
            if (_moduleDeclaration.DeclarationType == DeclarationType.ProceduralModule)
            {
                ((ProceduralModuleDeclaration)_moduleDeclaration).IsPrivateModule = true;
            }
            AddDeclaration(
                CreateDeclaration(
                    context.GetText(),
                    string.Empty,
                    Accessibility.Implicit,
                    DeclarationType.ModuleOption,
                    context,
                    context.GetSelection(),
                    false,
                    null,
                    null));
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var identifier = context.subroutineName();
            if (identifier == null)
            {
                return;
            }
            var name = context.subroutineName().GetText();
            var declaration = CreateDeclaration(
                name,
                null,
                accessibility,
                DeclarationType.Procedure,
                context,
                context.subroutineName().GetSelection(),
                false,
                null,
                null);
            AddDeclaration(declaration);
            SetCurrentScope(declaration, name);
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var identifier = context.functionName().identifier();
            if (identifier == null)
            {
                return;
            }
            var name = Identifier.GetName(identifier);

            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();
            var typeHint = Identifier.GetTypeHintValue(identifier);
            var isArray = asTypeClause != null && asTypeClause.type().LPAREN() != null;
            var declaration = CreateDeclaration(
                name,
                asTypeName,
                accessibility,
                DeclarationType.Function,
                context,
                context.functionName().identifier().GetSelection(),
                isArray,
                asTypeClause,
                typeHint);
            AddDeclaration(declaration);
            SetCurrentScope(declaration, name);
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var identifier = context.functionName().identifier();
            var name = Identifier.GetName(identifier);
            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();
            var typeHint = Identifier.GetTypeHintValue(identifier);
            var isArray = asTypeClause != null && asTypeClause.type().LPAREN() != null;
            var declaration = CreateDeclaration(
                name,
                asTypeName,
                accessibility,
                DeclarationType.PropertyGet,
                context,
                context.functionName().identifier().GetSelection(),
                isArray,
                asTypeClause,
                typeHint);

            AddDeclaration(declaration);
            SetCurrentScope(declaration, name);
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var identifier = context.subroutineName();
            if (identifier == null)
            {
                return;
            }
            var name = Identifier.GetName(identifier.identifier());
            var declaration = CreateDeclaration(
                name,
                null,
                accessibility,
                DeclarationType.PropertyLet,
                context,
                context.subroutineName().GetSelection(),
                false,
                null,
                null);
            AddDeclaration(declaration);
            SetCurrentScope(declaration, name);
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var identifier = context.subroutineName();
            if (identifier == null)
            {
                return;
            }
            var name = Identifier.GetName(identifier.identifier());

            var declaration = CreateDeclaration(
                name,
                null,
                accessibility,
                DeclarationType.PropertySet,
                context,
                context.subroutineName().GetSelection(),
                false,
                null,
                null);

            AddDeclaration(declaration);
            SetCurrentScope(declaration, name);
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterEventStmt(VBAParser.EventStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var identifier = context.identifier();
            if (identifier == null)
            {
                return;
            }
            var name = Identifier.GetName(identifier);

            var declaration = CreateDeclaration(
                name,
                null,
                accessibility,
                DeclarationType.Event,
                context,
                context.identifier().GetSelection(),
                false,
                null,
                null);

            AddDeclaration(declaration);
            SetCurrentScope(declaration, name);
        }

        public override void ExitEventStmt(VBAParser.EventStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterDeclareStmt(VBAParser.DeclareStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var identifier = context.identifier();
            if (identifier == null)
            {
                return;
            }
            var name = Identifier.GetName(identifier);
            var typeHint = Identifier.GetTypeHintValue(identifier);

            var hasReturnType = context.FUNCTION() != null;

            var asTypeClause = context.asTypeClause();
            var asTypeName = hasReturnType
                                ? asTypeClause == null
                                    ? Tokens.Variant
                                    : asTypeClause.type().GetText()
                                : null;
            var selection = identifier.GetSelection();

            var declarationType = hasReturnType
                ? DeclarationType.LibraryFunction
                : DeclarationType.LibraryProcedure;

            var declaration = CreateDeclaration(
                name,
                asTypeName,
                accessibility,
                declarationType,
                context,
                selection,
                false,
                asTypeClause,
                typeHint);

            AddDeclaration(declaration);
            SetCurrentScope(declaration, name); // treat like a procedure block, to correctly scope parameters.
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
                var identifier = argContext.unrestrictedIdentifier();
                string typeHint = Identifier.GetTypeHintValue(identifier);
                bool isArray = argContext.LPAREN() != null;
                AddDeclaration(
                    CreateDeclaration(
                        Identifier.GetName(identifier),
                        asTypeName,
                        Accessibility.Implicit,
                        DeclarationType.Parameter,
                        argContext,
                        identifier.GetSelection(),
                        isArray,
                        asTypeClause,
                        typeHint));
            }
        }

        public override void EnterStatementLabelDefinition(VBAParser.StatementLabelDefinitionContext context)
        {
            AddDeclaration(
                CreateDeclaration(
                    context.statementLabel().GetText(),
                    null,
                    Accessibility.Private,
                    DeclarationType.LineLabel,
                    context,
                    context.statementLabel().GetSelection(),
                    true,
                    null,
                    null));
        }

        public override void EnterVariableSubStmt(VBAParser.VariableSubStmtContext context)
        {
            var parent = (VBAParser.VariableStmtContext)context.Parent.Parent;
            var accessibility = GetMemberAccessibility(parent.visibility());
            var identifier = context.identifier();
            if (identifier == null)
            {
                return;
            }
            var name = Identifier.GetName(identifier);
            var typeHint = Identifier.GetTypeHintValue(identifier);
            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();
            var withEvents = parent.WITHEVENTS() != null;
            var isAutoObject = asTypeClause != null && asTypeClause.NEW() != null;
            bool isArray = context.LPAREN() != null;
            AddDeclaration(
                CreateDeclaration(
                    name,
                    asTypeName,
                    accessibility,
                    DeclarationType.Variable,
                    context,
                    context.identifier().GetSelection(),
                    isArray,
                    asTypeClause,
                    typeHint,
                    isAutoObject,
                    withEvents));
        }

        public override void EnterConstSubStmt(VBAParser.ConstSubStmtContext context)
        {
            var parent = (VBAParser.ConstStmtContext)context.Parent;
            var accessibility = GetMemberAccessibility(parent.visibility());

            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();
            var identifier = context.identifier();
            var typeHint = Identifier.GetTypeHintValue(identifier);
            var name = Identifier.GetName(identifier);
            var value = context.expression().GetText();
            var declaration = new ConstantDeclaration(
                new QualifiedMemberName(_qualifiedName, name),
                _parentDeclaration,
                _currentScope,
                asTypeName,
                asTypeClause,
                typeHint,
                accessibility,
                DeclarationType.Constant,
                value,
                context,
                identifier.GetSelection());

            AddDeclaration(declaration);
        }

        public override void EnterPublicTypeDeclaration(VBAParser.PublicTypeDeclarationContext context)
        {
            AddUdtDeclaration(context.udtDeclaration(), Accessibility.Public, context);
        }

        public override void ExitPublicTypeDeclaration(VBAParser.PublicTypeDeclarationContext context)
        {
            _parentDeclaration = _moduleDeclaration;
        }

        public override void EnterPrivateTypeDeclaration(VBAParser.PrivateTypeDeclarationContext context)
        {
            AddUdtDeclaration(context.udtDeclaration(), Accessibility.Private, context);
        }

        public override void ExitPrivateTypeDeclaration(VBAParser.PrivateTypeDeclarationContext context)
        {
            _parentDeclaration = _moduleDeclaration;
        }
        
        public void AddUdtDeclaration(VBAParser.UdtDeclarationContext udtDeclaration, Accessibility accessibility, ParserRuleContext context)
        {
            var identifier = Identifier.GetName(udtDeclaration.untypedIdentifier());
            var identifierSelection = Identifier.GetNameSelection(udtDeclaration.untypedIdentifier());
            var declaration = CreateDeclaration(
                identifier,
                null,
                accessibility,
                DeclarationType.UserDefinedType,
                context,
                identifierSelection,
                false,
                null,
                null);
            AddDeclaration(declaration);
            _parentDeclaration = declaration; // treat members as child declarations, but keep them scoped to module
        }

        public override void EnterUdtMember(VBAParser.UdtMemberContext context)
        {
            VBAParser.AsTypeClauseContext asTypeClause = null;
            bool isArray = false;
            string typeHint = null;
            string identifier;
            Selection identifierSelection;
            if (context.reservedNameMemberDeclaration() != null)
            {
                identifier = Identifier.GetName(context.reservedNameMemberDeclaration().unrestrictedIdentifier());
                identifierSelection = Identifier.GetNameSelection(context.reservedNameMemberDeclaration().unrestrictedIdentifier());
                asTypeClause = context.reservedNameMemberDeclaration().asTypeClause();
            }
            else
            {
                identifier = Identifier.GetName(context.untypedNameMemberDeclaration().untypedIdentifier());
                identifierSelection = Identifier.GetNameSelection(context.untypedNameMemberDeclaration().untypedIdentifier());
                asTypeClause = context.untypedNameMemberDeclaration().optionalArrayClause().asTypeClause();
                isArray = context.untypedNameMemberDeclaration().optionalArrayClause().arrayDim() != null;
            }
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();
            AddDeclaration(
                CreateDeclaration(
                    identifier,
                    asTypeName,
                    Accessibility.Implicit,
                    DeclarationType.UserDefinedTypeMember,
                    context,
                    identifierSelection,
                    isArray,
                    asTypeClause,
                    typeHint));
        }

        public override void EnterEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var identifier = context.identifier();
            if (identifier == null)
            {
                return;
            }
            var name = Identifier.GetName(identifier);

            var declaration = CreateDeclaration(
                name,
                "Long",
                accessibility,
                DeclarationType.Enumeration,
                context,
                context.identifier().GetSelection(),
                false,
                null,
                null);

            AddDeclaration(declaration);
            _parentDeclaration = declaration; // treat members as child declarations, but keep them scoped to module
        }

        public override void ExitEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            _parentDeclaration = _moduleDeclaration;
        }

        public override void EnterEnumerationStmt_Constant(VBAParser.EnumerationStmt_ConstantContext context)
        {
            AddDeclaration(CreateDeclaration(
                context.identifier().GetText(),
                "Long",
                Accessibility.Implicit,
                DeclarationType.EnumerationMember,
                context,
                context.identifier().GetSelection(),
                false,
                null,
                null));
        }

        private void AddDeclaration(Declaration declaration)
        {
            _createdDeclarations.Add(declaration);
        }
    }
}
