using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.DeclarationResolving
{
    public class DeclarationSymbolsListener : VBAParserBaseListener
    {
        private readonly QualifiedModuleName _qualifiedModuleName;
        private readonly Declaration _moduleDeclaration;

        private string _currentScope;
        private Declaration _currentScopeDeclaration;
        private Declaration _parentDeclaration;

        private readonly IDictionary<int, List<IParseTreeAnnotation>> _annotations;
        private readonly LogicalLineStore _logicalLines;
        private readonly IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> _attributes;
        private readonly IDictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext> _membersAllowingAttributes;

        private readonly List<Declaration> _createdDeclarations = new List<Declaration>();
        public IReadOnlyList<Declaration> CreatedDeclarations => _createdDeclarations;

        public DeclarationSymbolsListener(Declaration moduleDeclaration,
            IDictionary<int, List<IParseTreeAnnotation>> annotations,
            LogicalLineStore logicalLines,
            IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> attributes,
            IDictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext>
                membersAllowingAttributes)
        {
            _moduleDeclaration = moduleDeclaration;
            _qualifiedModuleName = moduleDeclaration.QualifiedModuleName;
            _annotations = annotations;
            _logicalLines = logicalLines;
            _attributes = attributes;
            _membersAllowingAttributes = membersAllowingAttributes;

            SetCurrentScope();
        }

        private IEnumerable<IParseTreeAnnotation> FindMemberAnnotations(int firstMemberLine)
        {
            return FindAnnotations(firstMemberLine, AnnotationTarget.Member);
        }

        private IEnumerable<IParseTreeAnnotation> FindAnnotations(int firstLine, AnnotationTarget requiredTarget)
        {
            if (_annotations == null)
            {
                return null;
            }

            var firstLineOfLogicalLine = _logicalLines.StartOfContainingLogicalLine(firstLine);
            if (!firstLineOfLogicalLine.HasValue)
            {
                return null;
            }

            if (_annotations.TryGetValue(firstLineOfLogicalLine.Value, out var scopedAnnotations))
            {
                return scopedAnnotations.Where(annotation => annotation.Annotation.Target.HasFlag(requiredTarget));
            }

            return Enumerable.Empty<IParseTreeAnnotation>();
        }

        private IEnumerable<IParseTreeAnnotation> FindVariableAnnotations(int firstVariableLine)
        {
            return FindAnnotations(firstVariableLine, AnnotationTarget.Variable);
        }

        private IEnumerable<IParseTreeAnnotation> FindGeneralAnnotations(int firstLine)
        {
            return FindAnnotations(firstLine, AnnotationTarget.General);
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

                var isByRef = argContext.BYREF() != null || argContext.BYVAL() == null;
                var isParamArray = argContext.PARAMARRAY() != null;

                result = new ParameterDeclaration(
                    new QualifiedMemberName(_qualifiedModuleName, identifierName),
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
                if (_parentDeclaration is IParameterizedDeclaration declaration)
                {
                    declaration.AddParameter((ParameterDeclaration)result);
                }
            }
            else
            {
                var key = (identifierName, declarationType);
                _attributes.TryGetValue(key, out var attributes);
                _membersAllowingAttributes.TryGetValue(key, out var attributesPassContext);

                switch (declarationType)
                {
                    case DeclarationType.Procedure:
                        result = new SubroutineDeclaration(
                            new QualifiedMemberName(_qualifiedModuleName, identifierName), 
                            _parentDeclaration, 
                            _currentScopeDeclaration, 
                            asTypeName, 
                            accessibility, 
                            context,
                            attributesPassContext,
                            selection, 
                            true,
                            FindMemberAnnotations(selection.StartLine), 
                            attributes);
                        break;
                    case DeclarationType.Function:
                        result = new FunctionDeclaration(
                            new QualifiedMemberName(_qualifiedModuleName, identifierName),
                            _parentDeclaration,
                            _currentScopeDeclaration,
                            asTypeName,
                            asTypeContext,
                            typeHint,
                            accessibility,
                            context,
                            attributesPassContext,
                            selection,
                            isArray,
                            true,
                            FindMemberAnnotations(selection.StartLine),
                            attributes);
                        break;
                    case DeclarationType.Event:
                        result = new EventDeclaration(
                            new QualifiedMemberName(_qualifiedModuleName, identifierName),
                            _parentDeclaration,
                            _currentScopeDeclaration,
                            asTypeName,
                            asTypeContext,
                            typeHint,
                            accessibility,
                            context,
                            selection,
                            isArray,
                            true,
                            FindGeneralAnnotations(selection.StartLine),
                            attributes);
                        break;
                    case DeclarationType.LibraryProcedure:
                    case DeclarationType.LibraryFunction:
                        result = new ExternalProcedureDeclaration(
                            new QualifiedMemberName(_qualifiedModuleName, identifierName), 
                            _parentDeclaration, 
                            _currentScopeDeclaration, 
                            declarationType, 
                            asTypeName, 
                            asTypeContext, 
                            accessibility, 
                            context,
                            attributesPassContext,
                            selection, 
                            true,
                            FindMemberAnnotations(selection.StartLine),
                            attributes);
                        break;
                    case DeclarationType.PropertyGet:
                        result = new PropertyGetDeclaration(
                            new QualifiedMemberName(_qualifiedModuleName, identifierName),
                            _parentDeclaration,
                            _currentScopeDeclaration,
                            asTypeName,
                            asTypeContext,
                            typeHint,
                            accessibility,
                            context,
                            attributesPassContext,
                            selection,
                            isArray,
                            true,
                            FindMemberAnnotations(selection.StartLine),
                            attributes);
                        break;
                    case DeclarationType.PropertySet:
                        result = new PropertySetDeclaration(
                            new QualifiedMemberName(_qualifiedModuleName, identifierName), 
                            _parentDeclaration, 
                            _currentScopeDeclaration, 
                            asTypeName, 
                            accessibility, 
                            context,
                            attributesPassContext,
                            selection, 
                            true,
                            FindMemberAnnotations(selection.StartLine), 
                            attributes);
                        break;
                    case DeclarationType.PropertyLet:
                        result = new PropertyLetDeclaration(
                            new QualifiedMemberName(_qualifiedModuleName, identifierName), 
                            _parentDeclaration, 
                            _currentScopeDeclaration, 
                            asTypeName, 
                            accessibility, 
                            context,
                            attributesPassContext,
                            selection, 
                            true,
                            FindMemberAnnotations(selection.StartLine), 
                            attributes);
                        break;
                    case DeclarationType.EnumerationMember:
                        result = new ValuedDeclaration(
                            new QualifiedMemberName(_qualifiedModuleName, identifierName),
                            _parentDeclaration, 
                            _currentScope, 
                            asTypeName, 
                            asTypeContext, 
                            typeHint,
                            FindVariableAnnotations(selection.StartLine),
                            accessibility, 
                            declarationType,
                            (context as VBAParser.EnumerationStmt_ConstantContext)?.expression()?.GetText() ?? string.Empty,
                            context,
                            selection);
                        break;
                    case DeclarationType.Variable:
                        result = new VariableDeclaration(
                            new QualifiedMemberName(_qualifiedModuleName, identifierName),
                            _parentDeclaration,
                            _currentScopeDeclaration,
                            asTypeName,
                            typeHint,
                            selfAssigned,
                            withEvents,
                            accessibility,
                            context,
                            attributesPassContext,
                            selection,
                            isArray,
                            asTypeContext,
                            FindVariableAnnotations(selection.StartLine),
                            attributes);
                        break;
                    default:
                        result = new Declaration(
                            new QualifiedMemberName(_qualifiedModuleName, identifierName),
                            _parentDeclaration,
                            _currentScopeDeclaration,
                            asTypeName,
                            typeHint,
                            selfAssigned,
                            withEvents,
                            accessibility,
                            declarationType,
                            context,
                            attributesPassContext,
                            selection,
                            isArray,
                            asTypeContext,
                            true,
                            FindGeneralAnnotations(selection.StartLine),
                            attributes);
                        break;
                }
                if (_parentDeclaration is ClassModuleDeclaration classParent && 
                    ((result as ICanBeDefaultMember)?.IsDefaultMember ?? false))
                {
                    classParent.DefaultMember = result;
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
            _currentScope = _qualifiedModuleName.ToString();
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
            _currentScope = _qualifiedModuleName + "." + name;
            _currentScopeDeclaration = procedureDeclaration;
            _parentDeclaration = procedureDeclaration;
        }

        public override void EnterImplementsStmt(VBAParser.ImplementsStmtContext context)
        {
            // The expression will be later resolved to the actual declaration. Have to split the work up because we have to gather/create all declarations first.
            ((ClassModuleDeclaration)_moduleDeclaration).AddSupertypeName(context.expression().GetText());
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
            var typeHint = Identifier.GetTypeHintValue(identifier);
            var asTypeClause = context.asTypeClause();
            var asTypeName = typeHint == null
                ? asTypeClause == null
                    ? Tokens.Variant
                    : asTypeClause.type().GetText()
                : SymbolList.TypeHintToTypeName[typeHint];
            var isArray = asTypeName.EndsWith("()");
            var actualAsTypeName = isArray && asTypeName.EndsWith("()")
                ? asTypeName.Substring(0, asTypeName.Length - 2)
                : asTypeName;
            var declaration = CreateDeclaration(
                name,
                actualAsTypeName,
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
            var typeHint = Identifier.GetTypeHintValue(identifier);
            var asTypeClause = context.asTypeClause();
            var asTypeName = typeHint == null
                ? asTypeClause == null
                    ? Tokens.Variant
                    : asTypeClause.type().GetText()
                : SymbolList.TypeHintToTypeName[typeHint];
            var isArray = asTypeClause != null && asTypeClause.type().LPAREN() != null;
            var actualAsTypeName = isArray && asTypeName.EndsWith("()")
                ? asTypeName.Substring(0, asTypeName.Length - 2)
                : asTypeName;
            var declaration = CreateDeclaration(
                name,
                actualAsTypeName,
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
                var identifier = argContext.unrestrictedIdentifier();
                string typeHint = Identifier.GetTypeHintValue(identifier);
                var asTypeClause = argContext.asTypeClause();
                var asTypeName = typeHint == null
                    ? asTypeClause == null
                        ? Tokens.Variant
                        : asTypeClause.type().GetText()
                    : SymbolList.TypeHintToTypeName[typeHint];
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
            if (context.combinedLabels() != null)
            {
                var combinedLabel = context.combinedLabels();
                AddIdentifierStatementLabelDeclaration(combinedLabel.identifierStatementLabel());
                AddLineNumberLabelDeclaration(combinedLabel.lineNumberLabel());
            }
            else if (context.identifierStatementLabel() != null)
            {
                AddIdentifierStatementLabelDeclaration(context.identifierStatementLabel());
            }
            else
            {
                AddLineNumberLabelDeclaration(context.standaloneLineNumberLabel().lineNumberLabel());
            }
        }

        private void AddIdentifierStatementLabelDeclaration(VBAParser.IdentifierStatementLabelContext context)
        {
            var statementText = context.legalLabelIdentifier().GetText();
            var statementSelection = context.legalLabelIdentifier().GetSelection();

            AddDeclaration(
                CreateDeclaration(
                    statementText,
                    null,
                    Accessibility.Private,
                    DeclarationType.LineLabel,
                    context,
                    statementSelection,
                    false,
                    null,
                    null));
        }

        private void AddLineNumberLabelDeclaration(VBAParser.LineNumberLabelContext context)
        {
            var statementText = context.GetText().Trim();
            var statementSelection = context.GetSelection();

            AddDeclaration(
                CreateDeclaration(
                    statementText,
                    null,
                    Accessibility.Private,
                    DeclarationType.LineLabel,
                    context,
                    statementSelection,
                    false,
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
            var asTypeName = typeHint == null
                ? asTypeClause == null
                    ? Tokens.Variant
                    : asTypeClause.type().GetText()
                : SymbolList.TypeHintToTypeName[typeHint];
            var withEvents = context.WITHEVENTS() != null;
            var isAutoObject = asTypeClause != null && asTypeClause.NEW() != null;
            bool isArray = context.arrayDim() != null;
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

            var identifier = context.identifier();
            var typeHint = Identifier.GetTypeHintValue(identifier);
            var asTypeClause = context.asTypeClause();
            var asTypeName = typeHint == null
                ? asTypeClause == null
                    ? Tokens.Variant
                    : asTypeClause.type().GetText()
                : SymbolList.TypeHintToTypeName[typeHint];
            var name = Identifier.GetName(identifier);
            var value = context.expression().GetText();
            var constStmt = (VBAParser.ConstStmtContext)context.Parent;

            var key = (name, DeclarationType.Constant);
            _attributes.TryGetValue(key, out var attributes);
            _membersAllowingAttributes.TryGetValue(key, out var attributesPassContext);

            var declaration = new ValuedDeclaration(
                new QualifiedMemberName(_qualifiedModuleName, name),
                _parentDeclaration,
                _currentScope,
                asTypeName,
                asTypeClause,
                typeHint,
                FindVariableAnnotations(constStmt.Start.Line),
                accessibility,
                DeclarationType.Constant,
                value,
                context,
                identifier.GetSelection(),
                attributesPassContext: attributesPassContext,
                attributes: attributes);

            AddDeclaration(declaration);
        }

        public override void EnterUdtDeclaration(VBAParser.UdtDeclarationContext context)
        {
            AddUdtDeclaration(context);
        }

        public override void ExitUdtDeclaration(VBAParser.UdtDeclarationContext context)
        {
            _parentDeclaration = _moduleDeclaration;
        }

        private void AddUdtDeclaration(VBAParser.UdtDeclarationContext context)
        {
            var identifier = Identifier.GetName(context.untypedIdentifier());
            var identifierSelection = Identifier.GetNameSelection(context.untypedIdentifier());
            var accessibility = GetAccessibility(context.visibility());
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

            Accessibility GetAccessibility(VBAParser.VisibilityContext visibilityContext)
            {
                if (visibilityContext == null)
                {
                    return Accessibility.Implicit;
                }

                if (visibilityContext.PUBLIC() != null)
                {
                    return Accessibility.Public;
                }

                return Accessibility.Private;
            }
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
                WithBracketsRemoved(context.identifier().GetText()),
                "Long",
                Accessibility.Implicit,
                DeclarationType.EnumerationMember,
                context,
                context.identifier().GetSelection(),
                false,
                null,
                null));
        }

        private static string WithBracketsRemoved(string enumElementName)
        {
            if (enumElementName.StartsWith("[") && enumElementName.EndsWith("]"))
            {
                return enumElementName.Substring(1, enumElementName.Length - 2);
            }

            return enumElementName;
        }

        public override void EnterOptionPrivateModuleStmt(VBAParser.OptionPrivateModuleStmtContext context)
        {
            ((ProceduralModuleDeclaration)_moduleDeclaration).IsPrivateModule = true;
        }

        private void AddDeclaration(Declaration declaration)
        {
            _createdDeclarations.Add(declaration);
        }
    }
}
