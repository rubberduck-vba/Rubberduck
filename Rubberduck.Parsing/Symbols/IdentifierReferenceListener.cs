using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReferenceListener : VBABaseListener
    {
        private enum ContextAccessorType
        {
            GetValueOrReference,
            AssignValue,
            AssignReference
        }

        private readonly Declarations _declarations;
        private readonly QualifiedModuleName _qualifiedModuleName;

        private readonly HashSet<DeclarationType> _moduleTypes;
        private readonly HashSet<DeclarationType> _memberTypes;

        private readonly Stack<Declaration> _withBlockQualifiers;

        public IdentifierReferenceListener(QualifiedModuleName qualifiedModuleName, Declarations declarations)
        {
            _qualifiedModuleName = qualifiedModuleName;
            _declarations = declarations;

            _moduleTypes = new HashSet<DeclarationType>(new[]
            {
                DeclarationType.Module, 
                DeclarationType.Class
            });

            _memberTypes = new HashSet<DeclarationType>(new[]
            {
                DeclarationType.Function, 
                DeclarationType.Procedure, 
                DeclarationType.PropertyGet, 
                DeclarationType.PropertyLet, 
                DeclarationType.PropertySet
            });

            _withBlockQualifiers = new Stack<Declaration>();

            SetCurrentScope();
        }

        private Declaration _currentScope;

        /// <summary>
        /// Sets the current scope to module-level.
        /// </summary>
        private void SetCurrentScope()
        {
            _currentScope = _declarations.Items.Single(item =>
                _moduleTypes.Contains(item.DeclarationType)
                && item.Project == _qualifiedModuleName.Project
                && item.ComponentName == _qualifiedModuleName.ComponentName);
        }

        /// <summary>
        /// Sets the current scope to the specified member.
        /// </summary>
        /// <param name="memberName">The name of the member.</param>
        /// <param name="accessor">For properties, specifies the accessor type.</param>
        private void SetCurrentScope(string memberName, DeclarationType? accessor = null)
        {
            _currentScope = _declarations.Items.Single(item =>
                _memberTypes.Contains(item.DeclarationType)
                && (!accessor.HasValue || item.DeclarationType == accessor.Value)
                && item.Project == _qualifiedModuleName.Project
                && item.ComponentName == _qualifiedModuleName.ComponentName
                && item.IdentifierName == memberName);
        }

        #region IVBAListener scoping overrides

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText());
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText());
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyGet);
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyLet);
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertySet);
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterWithStmt(VBAParser.WithStmtContext context)
        {
            Declaration qualifier;
            if (context.NEW() == null)
            {
                // with block is using an identifier declared elsewhere.
                qualifier = ResolveType(context.implicitCallStmt_InStmt());
            }
            else
            {
                // with block is using an anonymous declaration.
                // i.e. object variable reference is held by the with block itself.
                var typeContext = context.type();
                var baseTypeContext = typeContext.baseType().COLLECTION();
                if (baseTypeContext != null)
                {
                    // object variable is a built-in Collection class instance
                    qualifier = _declarations.Items.Single(item => item.IsBuiltIn
                        && item.IdentifierName == baseTypeContext.GetText());
                }
                else
                {
                    qualifier = ResolveType(typeContext.complexType());
                }
            }

            _withBlockQualifiers.Push(qualifier); // note: pushes null if unresolved
        }

        public override void ExitWithStmt(VBAParser.WithStmtContext context)
        {
            _withBlockQualifiers.Pop();
        }

        #endregion

        #region ResolveType() overloads

        private Declaration ResolveType(VBAParser.ComplexTypeContext context)
        {
            if (context == null)
            {
                return null;
            }

            var identifiers = context.ambiguousIdentifier();

            // VBA doesn't support namespaces.
            // A "ComplexType" is therefore only ever as "complex" as [Library].[Type].
            var identifier = identifiers.Last();
            var library = identifiers.Count > 1
                ? identifiers[0]
                : null;

            var libraryName = library == null
                ? _qualifiedModuleName.ProjectName
                : library.GetText();

            // note: inter-project references won't work, but we can qualify VbaStandardLib types:
            if (libraryName == _qualifiedModuleName.ProjectName || libraryName == "VBA")
            {
                return _declarations[identifier.GetText()].SingleOrDefault(item =>
                    item.ProjectName == libraryName
                    && item.DeclarationType == DeclarationType.Class);
            }

            return null;
        }

        private Declaration ResolveType(VBAParser.ImplicitCallStmt_InStmtContext context, Declaration localScope = null)
        {
            if (context == null)
            {
                return null;
            }

            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var type = Resolve(context.iCS_S_VariableOrProcedureCall(), localScope)
                ?? Resolve(context.iCS_S_ProcedureOrArrayCall(), localScope)
                ?? Resolve(context.iCS_S_MembersCall(), localScope)
                ?? Resolve(context.iCS_S_DictionaryCall(), localScope, ContextAccessorType.GetValueOrReference, context.iCS_S_DictionaryCall().dictionaryCallStmt());

            return ResolveType(type);
        }

        private Declaration ResolveType(Declaration parent)
        {
            if (parent == null || parent.AsTypeName == null)
            {
                return null;
            }

            // look in current project first.
            var result = _declarations[parent.AsTypeName].SingleOrDefault(item =>
                _moduleTypes.Contains(item.DeclarationType)
                && item.ProjectName == _currentScope.ProjectName);

            if (result == null)
            {
                // look in all projects (including VbaStdLib library.
                result = _declarations[parent.AsTypeName].SingleOrDefault(item =>
                    _moduleTypes.Contains(item.DeclarationType));
            }

            return result;
        }

        #endregion

        #region Resolve() overloads

        private Declaration Resolve(ParserRuleContext callSiteContext, Declaration localScope, ContextAccessorType accessorType = ContextAccessorType.GetValueOrReference, VBAParser.DictionaryCallStmtContext fieldCall = null, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (callSiteContext == null)
            {
                return null;
            }

            /* VBA allows ambiguous identifiers; if foo is declared at both
             * local and module scope, local scope takes precedence.
             * Identifier reference resolution should therefore start search for 
             * declarations in this order:
             *  1. Local scope (variable)
             *  2a. Module scope (variable)
             *  2b. Module scope (procedure)
             *  3a. Project/Global scope (variable)
             *  3b. Project/Global scope (procedure)
             *  4a. Global (references) scope (variable)*
             *  4b. Global (references) scope (procedure)*
             *  
             *  *project references aren't accounted for... yet.
             */

            if (localScope == null)
            {
                localScope = _currentScope;
            }

            if (_withBlockQualifiers.Any())
            {
                localScope = _withBlockQualifiers.Peek();
            }

            var identifierName = callSiteContext.GetText();
            var callee = FindLocalScopeDeclaration(identifierName, localScope)
                         ?? FindModuleScopeDeclaration(identifierName, localScope)
                         ?? FindModuleScopeProcedure(identifierName, localScope, accessorType)
                         ?? FindProjectScopeDeclaration(identifierName);

            if (callee == null)
            {
                return null;
            }

            var reference = CreateReference(callSiteContext, callee, isAssignmentTarget, hasExplicitLetStatement);
            callee.AddReference(reference);

            if (fieldCall != null)
            {
                return Resolve(fieldCall, callee);
            }

            return callee;
        }

        private Declaration Resolve(VBAParser.ICS_S_VariableOrProcedureCallContext context, Declaration localScope, ContextAccessorType accessorType = ContextAccessorType.GetValueOrReference, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (context == null)
            {
                return null;
            }

            _isResolving = true;

            var identifierContext = context.ambiguousIdentifier();
            var fieldCall = context.dictionaryCallStmt();

            return Resolve(identifierContext, localScope, accessorType, fieldCall, hasExplicitLetStatement, isAssignmentTarget);
        }

        private Declaration Resolve(VBAParser.DictionaryCallStmtContext fieldCall, Declaration parent, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            var parentType = ResolveType(parent);
            var members = _declarations.FindMembers(parentType);
            var fieldName = fieldCall.ambiguousIdentifier().GetText();

            var result = members.SingleOrDefault(member => member.IdentifierName == fieldName);
            if (result == null)
            {
                return null;
            }

            var identifierContext = fieldCall.ambiguousIdentifier();
            var reference = CreateReference(identifierContext, result, isAssignmentTarget, hasExplicitLetStatement);
            result.AddReference(reference);

            return result;
        }

        private Declaration Resolve(VBAParser.ICS_S_ProcedureOrArrayCallContext context, Declaration localScope, ContextAccessorType accessorType = ContextAccessorType.GetValueOrReference, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (context == null)
            {
                return null;
            }

            _isResolving = true;

            var identifierContext = context.ambiguousIdentifier();
            var fieldCall = context.dictionaryCallStmt();
            // todo: understand WTF [baseType] is doing in that grammar rule...

            return Resolve(identifierContext, localScope, accessorType, fieldCall, hasExplicitLetStatement, isAssignmentTarget);
        }

        private Declaration Resolve(VBAParser.ICS_S_MembersCallContext context, ContextAccessorType accessorType, Declaration localScope = null, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (context == null)
            {
                return null;
            }

            _isResolving = true;

            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var parent = Resolve(context.iCS_S_ProcedureOrArrayCall(), localScope, accessorType, hasExplicitLetStatement, isAssignmentTarget)
                      ?? Resolve(context.iCS_S_VariableOrProcedureCall(), localScope, accessorType, hasExplicitLetStatement, isAssignmentTarget);

            if (parent == null)
            {
                return null;
            }

            var parentReference = CreateReference(parent.Context, parent);
            parent.AddReference(parentReference);

            var chainedCalls = context.iCS_S_MemberCall();
            foreach (var memberCall in chainedCalls)
            {
                var member = Resolve(memberCall.iCS_S_ProcedureOrArrayCall(), parent, accessorType, hasExplicitLetStatement, isAssignmentTarget)
                          ?? Resolve(memberCall.iCS_S_VariableOrProcedureCall(), parent, accessorType, hasExplicitLetStatement, isAssignmentTarget);

                if (member == null)
                {
                    return null;
                }

                parent = member;
            }

            var fieldCall = context.dictionaryCallStmt();
            if (fieldCall == null)
            {
                return parent;
            }

            return Resolve(fieldCall, parent, hasExplicitLetStatement, isAssignmentTarget);
        }

        private void Resolve(VBAParser.ICS_B_ProcedureCallContext context)
        {
            if (context == null)
            {
                return;
            }

            _isResolving = true;

            var identifierContext = context.certainIdentifier();
            var callee = Resolve(identifierContext, _currentScope);
            if (callee == null)
            {
                return;
            }

            var reference = CreateReference(identifierContext, callee);
            callee.AddReference(reference);
        }

        private Declaration Resolve(VBAParser.ImplicitCallStmt_InStmtContext callSiteContext, Declaration localScope, ContextAccessorType accessorType, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            var dictionaryCall = callSiteContext.iCS_S_DictionaryCall();
            var fieldCall = dictionaryCall == null ? null : dictionaryCall.dictionaryCallStmt();

            return Resolve(callSiteContext.iCS_S_VariableOrProcedureCall(), localScope, accessorType, hasExplicitLetStatement, isAssignmentTarget)
                ?? Resolve(callSiteContext.iCS_S_ProcedureOrArrayCall(), localScope, accessorType, hasExplicitLetStatement, isAssignmentTarget)
                ?? Resolve(callSiteContext.iCS_S_MembersCall(), accessorType, localScope, hasExplicitLetStatement, isAssignmentTarget)
                ?? Resolve(callSiteContext.iCS_S_DictionaryCall(), localScope, accessorType, fieldCall, hasExplicitLetStatement, isAssignmentTarget);
        }

        #endregion

        #region FindScopedDeclaration methods

        /// <summary>
        /// Finds a declaration located in the local scope.
        /// </summary>
        /// <param name="identifierName">The name of the identifier to find.</param>
        /// <param name="localScope">The scope considered local.</param>
        /// <returns></returns>
        private Declaration FindLocalScopeDeclaration(string identifierName, Declaration localScope = null)
        {
            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var parent = _declarations[identifierName].SingleOrDefault(item =>
                item.ParentScope == localScope.Scope);
            return parent;
        }

        /// <summary>
        /// Finds a module-scope variable in the specified scope.
        /// </summary>
        /// <param name="identifierName">The name of the identifier to find.</param>
        /// <param name="localScope">The scope considered local.</param>
        /// <returns></returns>
        private Declaration FindModuleScopeDeclaration(string identifierName, Declaration localScope = null)
        {
            if (localScope == null)
            {
                localScope = _currentScope;
            }

            return _declarations[identifierName].SingleOrDefault(item =>
                item.ParentScope == localScope.ParentScope
                && !item.DeclarationType.HasFlag(DeclarationType.Member));
        }

        /// <summary>
        /// Finds a procedure declaration in the specified scope.
        /// </summary>
        /// <param name="identifierName">The name of the identifier to find.</param>
        /// <param name="localScope">The scope considered local.</param>
        /// <param name="accessorType">Disambiguates <see cref="DeclarationType.PropertyLet"/> and <see cref="DeclarationType.PropertySet"/> accessors.</param>
        private Declaration FindModuleScopeProcedure(string identifierName, Declaration localScope, ContextAccessorType accessorType)
        {
            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var result = _declarations[identifierName].Where(item =>
                item.ParentScope == localScope.ParentScope
                && (item.DeclarationType == DeclarationType.Function
                 || item.DeclarationType == DeclarationType.Procedure
                 || (accessorType == ContextAccessorType.GetValueOrReference && item.DeclarationType == DeclarationType.PropertyGet)
                 || ((accessorType == ContextAccessorType.AssignValue && item.DeclarationType.HasFlag(DeclarationType.PropertyLet) 
                    && (localScope.DeclarationType == item.DeclarationType) || localScope.ParentScope != item.ParentScope)
                  || (accessorType == ContextAccessorType.AssignReference && item.DeclarationType.HasFlag(DeclarationType.PropertySet)
                    && localScope.DeclarationType == item.DeclarationType) || localScope.ParentScope != item.ParentScope)))
                  .ToList();

            return result.SingleOrDefault();
        }

        /// <summary>
        /// Finds a global (project) scope declaration for an unqualified (unambiguous) call.
        /// </summary>
        /// <param name="identifierName"></param>
        /// <returns></returns>
        private Declaration FindProjectScopeDeclaration(string identifierName)
        {
            // assume unqualified variable call, i.e. unique declaration.
            return _declarations[identifierName].SingleOrDefault(item =>
                item.Accessibility == Accessibility.Public
                || item.Accessibility == Accessibility.Global);
        }

        /// <summary>
        /// Finds a global (project) scope declaration for a qualified call.
        /// </summary>
        /// <param name="identifierName"></param>
        /// <param name="moduleName"></param>
        /// <returns></returns>
        private Declaration FindProjectScopeDeclaration(string identifierName, string moduleName)
        {
            return _declarations[identifierName].SingleOrDefault(item =>
                item.ComponentName == moduleName &&
                (item.Accessibility == Accessibility.Public
                || item.Accessibility == Accessibility.Global));
        }

        #endregion

        #region IVBAListener overrides

        // avoid re-resolving identifiers
        private bool _isResolving;

        public override void EnterICS_B_ProcedureCall(VBAParser.ICS_B_ProcedureCallContext context)
        {
            if (_isResolving)
            {
                return;
            }

            Resolve(context);
            _isResolving = false;
        }

        public override void EnterICS_S_VariableOrProcedureCall(VBAParser.ICS_S_VariableOrProcedureCallContext context)
        {
            if (_isResolving)
            {
                return;
            }

            Resolve(context, _currentScope);
        }

        public override void EnterICS_S_ProcedureOrArrayCall(VBAParser.ICS_S_ProcedureOrArrayCallContext context)
        {
            if (_isResolving)
            {
                return;
            }

            Resolve(context, _currentScope);
        }

        public override void EnterICS_S_MembersCall(VBAParser.ICS_S_MembersCallContext context)
        {
            if (_isResolving)
            {
                return;
            }

            Resolve(context, _currentScope);
        }

        public override void EnterICS_S_DictionaryCall(VBAParser.ICS_S_DictionaryCallContext context)
        {
            if (_isResolving)
            {
                return;
            }

            Resolve(context, _currentScope);
        }

        public override void EnterLetStmt(VBAParser.LetStmtContext context)
        {
            var leftSide = context.implicitCallStmt_InStmt();
            var letStatement = context.LET();
            Resolve(leftSide, _currentScope, ContextAccessorType.AssignValue, letStatement != null, true);
        }

        public override void EnterSetStmt(VBAParser.SetStmtContext context)
        {
            var leftSide = context.implicitCallStmt_InStmt();
            Resolve(leftSide, _currentScope, ContextAccessorType.AssignReference, false, true);
        }

        #endregion

        private IdentifierReference CreateReference(ParserRuleContext callSiteContext, Declaration callee, bool isAssignmentTarget = false, bool hasExplicitLetStatement = false)
        {
            var name = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            return new IdentifierReference(_qualifiedModuleName, name, selection, callSiteContext, callee, isAssignmentTarget, hasExplicitLetStatement);
        }
    }

    //public class IdentifierReferenceListener : VBABaseListener
    //{
    //    private readonly Declarations _declarations;
    //    private readonly QualifiedModuleName _qualifiedName;

    //    private string _currentScope;
    //    private DeclarationType _currentScopeType;

    //    public IdentifierReferenceListener(VBComponentParseResult result, Declarations declarations)
    //        : this(result.QualifiedName, declarations)
    //    { }

    //    public IdentifierReferenceListener(QualifiedModuleName qualifiedName, Declarations declarations)
    //    {
    //        _qualifiedName = qualifiedName;
    //        _declarations = declarations;
    //        SetCurrentScope();
    //    }

    //    private string ModuleScope { get { return _qualifiedName.ToString(); } }

    //    /// <summary>
    //    /// Sets current scope to module-level.
    //    /// </summary>
    //    private void SetCurrentScope()
    //    {
    //        _currentScope = ModuleScope;
    //        _currentScopeType = _qualifiedName.Component.Type == vbext_ComponentType.vbext_ct_StdModule
    //            ? DeclarationType.Module
    //            : DeclarationType.Class;
    //    }

    //    /// <summary>
    //    /// Sets current scope to specified module member.
    //    /// </summary>
    //    private void SetCurrentScope(string name, DeclarationType scopeType)
    //    {
    //        _currentScope = _qualifiedName + "." + name;
    //        _currentScopeType = scopeType;
    //    }

    //    public override void EnterSubStmt(VBAParser.SubStmtContext context)
    //    {
    //        SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.Procedure);
    //    }

    //    public override void ExitSubStmt(VBAParser.SubStmtContext context)
    //    {
    //        SetCurrentScope();
    //    }

    //    public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
    //    {
    //        SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.Function);
    //    }

    //    public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
    //    {
    //        SetCurrentScope();
    //    }

    //    public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
    //    {
    //        SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyGet);
    //    }

    //    public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
    //    {
    //        SetCurrentScope();
    //    }

    //    public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
    //    {
    //        SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertyLet);
    //    }

    //    public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
    //    {
    //        SetCurrentScope();
    //    }

    //    public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
    //    {
    //        SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.PropertySet);
    //    }

    //    public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
    //    {
    //        SetCurrentScope();
    //    }

    //    public override void EnterLetStmt(VBAParser.LetStmtContext context)
    //    {
    //        var leftSide = context.implicitCallStmt_InStmt();
    //        var letStatement = context.LET();
    //        var target = FindAssignmentTarget(leftSide, DeclarationType.PropertyLet);
    //        if (target != null)
    //        {
    //            EnterIdentifier(target, target.GetSelection(), true, letStatement != null);
    //        }
    //    }

    //    public override void EnterSetStmt(VBAParser.SetStmtContext context)
    //    {
    //        var leftSide = context.implicitCallStmt_InStmt();
    //        var target = FindAssignmentTarget(leftSide, DeclarationType.PropertySet);
    //        if (target != null)
    //        {
    //            EnterIdentifier(target, target.GetSelection(), true);
    //        }
    //    }

    //    private VBAParser.AmbiguousIdentifierContext FindAssignmentTarget(VBAParser.ImplicitCallStmt_InStmtContext leftSide, DeclarationType accessorType)
    //    {
    //        VBAParser.AmbiguousIdentifierContext context;
    //        var call = Resolve(leftSide.iCS_S_ProcedureOrArrayCall(), out context, accessorType)
    //                   ?? Resolve(leftSide.iCS_S_VariableOrProcedureCall(), out context, accessorType)
    //                   ?? Resolve(leftSide.iCS_S_DictionaryCall(), out context, accessorType)
    //                   ?? Resolve(leftSide.iCS_S_MembersCall(), out context, accessorType);

    //        return context;
    //    }

    //    private VBAParser.AmbiguousIdentifierContext EnterDictionaryCall(VBAParser.DictionaryCallStmtContext dictionaryCall, VBAParser.AmbiguousIdentifierContext parentIdentifier = null, DeclarationType accessorType = DeclarationType.PropertyGet)
    //    {
    //        if (dictionaryCall == null)
    //        {
    //            return null;
    //        }

    //        if (parentIdentifier != null)
    //        {
    //            var isTarget = accessorType == DeclarationType.PropertyLet || accessorType == DeclarationType.PropertySet;
    //            if (!EnterIdentifier(parentIdentifier, parentIdentifier.GetSelection(), isTarget, accessorType:accessorType))
    //                // we're referencing "member" in "member!field"
    //            {
    //                return null;
    //            }
    //        }

    //        var identifier = dictionaryCall.ambiguousIdentifier();
    //        if (_declarations.Items.Any(item => item.IdentifierName == identifier.GetText()))
    //        {
    //            return identifier;
    //        }

    //        return null;
    //    }

    //    public override void EnterComplexType(VBAParser.ComplexTypeContext context)
    //    {
    //        var identifiers = context.ambiguousIdentifier();
    //        _skipIdentifiers = !identifiers.All(identifier => _declarations.Items.Any(declaration => declaration.IdentifierName == identifier.GetText()));
    //    }

    //    public override void ExitComplexType(VBAParser.ComplexTypeContext context)
    //    {
    //        _skipIdentifiers = false;
    //    }

    //    private bool _skipIdentifiers;
    //    public override void EnterAmbiguousIdentifier(VBAParser.AmbiguousIdentifierContext context)
    //    {
    //        if (_skipIdentifiers || IsDeclarativeContext(context))
    //        {
    //            return;
    //        }

    //        var selection = context.GetSelection();

    //        if (IsAssignmentContext(context))
    //        {
    //            EnterIdentifier(context, selection, true);
    //        }
    //        else
    //        {
    //            EnterIdentifier(context, selection);
    //        }
    //    }

    //    private Stack<Declaration> _parentMember;
    //    public override void EnterECS_MemberProcedureCall(VBAParser.ECS_MemberProcedureCallContext context)
    //    {
    //        var implicitCall = context.implicitCallStmt_InStmt();
    //        var member = Resolve(implicitCall.iCS_S_VariableOrProcedureCall())
    //                     ?? Resolve(implicitCall.iCS_S_ProcedureOrArrayCall())
    //                     ?? Resolve(implicitCall.iCS_S_DictionaryCall())
    //                     ?? Resolve(implicitCall.iCS_S_MembersCall());

    //        if (member == null && implicitCall.Start.Text == Tokens.Me)
    //        {
    //            member = _declarations[_qualifiedName.ComponentName].SingleOrDefault(item => item.DeclarationType == DeclarationType.Class);
    //        }
    //        if (member == null)
    //        {
    //            return;
    //        }

    //        EnterIdentifier(member.Context, member.Selection);

    //        var identifier = context.ambiguousIdentifier();
    //        EnterIdentifier(identifier, identifier.GetSelection());
    //    }

    //    public override void EnterECS_ProcedureCall(VBAParser.ECS_ProcedureCallContext context)
    //    {
    //        var identifier = context.ambiguousIdentifier();
    //        EnterIdentifier(identifier, identifier.GetSelection());
    //    }

    //    public override void EnterICS_S_MembersCall(VBAParser.ICS_S_MembersCallContext context)
    //    {
    //        var member = Resolve(context);
    //        if (member == null)
    //        {
    //            return;
    //        }

    //        EnterIdentifier(member.Context, member.Selection);
    //    }

    //    private bool IsAssignmentContext(ParserRuleContext context)
    //    {
    //        return context.Parent is VBAParser.ForNextStmtContext
    //               || context.Parent is VBAParser.ForEachStmtContext
    //               || context.Parent.Parent.Parent.Parent is VBAParser.LineInputStmtContext
    //               || context.Parent.Parent.Parent.Parent is VBAParser.InputStmtContext;
    //    }

    //    public override void EnterCertainIdentifier(VBAParser.CertainIdentifierContext context)
    //    {
    //        // skip declarations
    //        if (IsDeclarativeContext(context))
    //        {
    //            return;
    //        }

    //        var selection = context.GetSelection();
    //        EnterIdentifier(context, selection);
    //    }

    //    private bool EnterIdentifier(ParserRuleContext context, Selection selection, bool isAssignmentTarget = false, bool hasExplicitLetStatement = false, DeclarationType accessorType = DeclarationType.PropertyGet)
    //    {
    //        if (context == null)
    //        {
    //            return false;
    //        }

    //        var name = context.GetText();
    //        var matches = _declarations[name].Where(IsInScope);

    //        var declaration = GetClosestScopeDeclaration(matches, context, accessorType);
    //        if (declaration != null)
    //        {
    //            var reference = new IdentifierReference(_qualifiedName, name, selection, context, declaration, isAssignmentTarget, hasExplicitLetStatement);
    //            declaration.AddReference(reference); // doesn't re-add an existing one
    //            return true;

    //            // note: non-matching names are not necessarily undeclared identifiers, e.g. "String" in "Dim foo As String".
    //        }

    //        return false;
    //    }

    //    public override void EnterVsNew(VBAParser.VsNewContext context)
    //    {
    //        _skipIdentifiers = true;
    //        var identifiers = context.valueStmt().GetRuleContexts<VBAParser.ImplicitCallStmt_InStmtContext>();

    //        var lastIdentifier = identifiers.Last();
    //        var name = lastIdentifier.GetText();

    //        var matches = _declarations[name].Where(d => d.DeclarationType == DeclarationType.Class).ToList();
    //        var result = matches.Count <= 1 
    //            ? matches.SingleOrDefault()
    //            : GetClosestScopeDeclaration(matches, context, DeclarationType.Class);

    //        if (result == null)
    //        {
    //            return;
    //        }

    //        var reference = new IdentifierReference(_qualifiedName, result.IdentifierName, lastIdentifier.GetSelection(), context, result);
    //        result.AddReference(reference);
    //    }

    //    public override void ExitVsNew(VBAParser.VsNewContext context)
    //    {
    //        _skipIdentifiers = false;
    //    }

    //    private readonly Stack<Declaration> _withQualifiers = new Stack<Declaration>();
    //    public override void EnterWithStmt(VBAParser.WithStmtContext context)
    //    {
    //        var implicitCall = context.implicitCallStmt_InStmt();

    //        var call = Resolve(implicitCall.iCS_S_ProcedureOrArrayCall())
    //            ?? Resolve(implicitCall.iCS_S_VariableOrProcedureCall())
    //            ?? Resolve(implicitCall.iCS_S_DictionaryCall())
    //            ?? Resolve(implicitCall.iCS_S_MembersCall());

    //        _withQualifiers.Push(GetReturnType(call));            
    //    }

    //    private Declaration GetReturnType(Declaration call)
    //    {
    //        return call == null 
    //            ? null 
    //            : _declarations.Items.SingleOrDefault(item =>
    //            item.DeclarationType == DeclarationType.Class
    //            && item.Accessibility != Accessibility.Private
    //            && item.IdentifierName == call.AsTypeName);
    //    }

    //    public override void ExitWithStmt(VBAParser.WithStmtContext context)
    //    {
    //        _withQualifiers.Pop();
    //    }

    //    private Declaration Resolve(VBAParser.ICS_S_ProcedureOrArrayCallContext context, out VBAParser.AmbiguousIdentifierContext identifierContext, DeclarationType accessorType)
    //    {
    //        if (context == null)
    //        {
    //            identifierContext = null;
    //            return null;
    //        }

    //        var identifier = context.ambiguousIdentifier();
    //        var name = identifier.GetText();

    //        var procedure = FindProcedureDeclaration(name, identifier);
    //        var result = procedure ?? FindVariableDeclaration(name, identifier, accessorType);

    //        identifierContext = result == null 
    //            ? null 
    //            : result.Context == null 
    //                ? null 
    //                : ((dynamic) result.Context).ambiguousIdentifier();
    //        return result;
    //    }

    //    private Declaration Resolve(VBAParser.ICS_S_ProcedureOrArrayCallContext context)
    //    {
    //        VBAParser.AmbiguousIdentifierContext discarded;
    //        return Resolve(context, out discarded, DeclarationType.PropertyGet);
    //    }

    //    private Declaration Resolve(VBAParser.ICS_S_VariableOrProcedureCallContext context, out VBAParser.AmbiguousIdentifierContext identifierContext, DeclarationType accessorType)
    //    {
    //        if (context == null)
    //        {
    //            identifierContext = null;
    //            return null;
    //        }

    //        var identifier = context.ambiguousIdentifier();
    //        var name = identifier.GetText();

    //        var procedure = FindProcedureDeclaration(name, identifier, accessorType);
    //        var result = procedure ?? FindVariableDeclaration(name, identifier, accessorType);

    //        identifierContext = result == null 
    //            ? null 
    //            : result.Context == null 
    //                ? null 
    //                : ((dynamic) result.Context).ambiguousIdentifier();
    //        return result;
    //    }

    //    private Declaration Resolve(VBAParser.ICS_S_VariableOrProcedureCallContext context)
    //    {
    //        VBAParser.AmbiguousIdentifierContext discarded;
    //        return Resolve(context, out discarded, DeclarationType.PropertyGet);
    //    }

    //    private Declaration Resolve(VBAParser.ICS_S_DictionaryCallContext context, out VBAParser.AmbiguousIdentifierContext identifierContext, DeclarationType accessorType, VBAParser.AmbiguousIdentifierContext parentIdentifier = null)
    //    {
    //        if (context == null)
    //        {
    //            identifierContext = null;
    //            return null;
    //        }

    //        var identifier = EnterDictionaryCall(context.dictionaryCallStmt(), parentIdentifier, accessorType);
    //        var name = identifier.GetText();

    //        var result = FindVariableDeclaration(name, identifier, accessorType);

    //        identifierContext = result == null 
    //            ? null 
    //            : result.Context == null 
    //                ? null 
    //                : ((dynamic) result.Context).ambiguousIdentifier();
    //        return result;
    //    }

    //    private Declaration Resolve(VBAParser.ICS_S_DictionaryCallContext context, VBAParser.AmbiguousIdentifierContext parentIdentifier = null)
    //    {
    //        VBAParser.AmbiguousIdentifierContext discarded;
    //        return Resolve(context, out discarded, DeclarationType.PropertyGet, parentIdentifier);
    //    }

    //    private Declaration Resolve(VBAParser.ICS_S_MembersCallContext context, out VBAParser.AmbiguousIdentifierContext identifierContext, DeclarationType accessorType)
    //    {
    //        if (context == null)
    //        {
    //            identifierContext = null;
    //            return null;
    //        }

    //        var members = context.iCS_S_MemberCall();
    //        for (var index = 0; index < members.Count; index++)
    //        {
    //            var member = members[index];
    //            if (index < members.Count - 1)
    //            {
    //                var parent = Resolve(member.iCS_S_ProcedureOrArrayCall())
    //                                     ?? Resolve(member.iCS_S_VariableOrProcedureCall());

    //                if (parent == null)
    //                {
    //                    // return early if we can't resolve the whole member chain
    //                    identifierContext = null;
    //                    return null;
    //                }
    //            }
    //            else
    //            {
    //                var result = Resolve(member.iCS_S_ProcedureOrArrayCall())
    //                             ?? Resolve(member.iCS_S_VariableOrProcedureCall());

    //                identifierContext = result == null 
    //                    ? null 
    //                    : result.Context == null 
    //                        ? null 
    //                        : ((dynamic) result.Context).ambiguousIdentifier();
    //                return result;
    //            }
    //        }

    //        identifierContext = null;
    //        return null;
    //    }

    //    private Declaration Resolve(VBAParser.ICS_S_MembersCallContext context)
    //    {
    //        VBAParser.AmbiguousIdentifierContext discarded;
    //        return Resolve(context, out discarded, DeclarationType.PropertyGet);
    //    }

    //    private Declaration Resolve(VBAParser.ICS_B_MemberProcedureCallContext context)
    //    {
    //        Declaration type;
    //        IEnumerable<Declaration> members;
    //        var name = context.ambiguousIdentifier().GetText();

    //        var parent = context.implicitCallStmt_InStmt();
    //        if (parent == null && _withQualifiers.Any())
    //        {
    //            type = _withQualifiers.Pop();
    //            members = _declarations.FindMembers(type);
    //            return members.SingleOrDefault(m => m.IdentifierName == name);
    //        }
    //        if (parent == null)
    //        {
    //            return null; // bug in grammar..
    //        }

    //        var parentCall = Resolve(parent.iCS_S_VariableOrProcedureCall())
    //                         ?? Resolve(parent.iCS_S_ProcedureOrArrayCall())
    //                         ?? Resolve(parent.iCS_S_DictionaryCall())
    //                         ?? Resolve(parent.iCS_S_MembersCall());

    //        if (parentCall == null)
    //        {
    //            return parent.Start.Text == Tokens.Me 
    //                ? _declarations[_qualifiedName.ComponentName].SingleOrDefault(item => item.DeclarationType == DeclarationType.Class)
    //                : null;
    //        }

    //        type = _declarations[parentCall.AsTypeName].SingleOrDefault(item =>
    //            item.DeclarationType == DeclarationType.Class
    //            || item.DeclarationType == DeclarationType.UserDefinedType);

    //        if (type == null)
    //        {
    //            return null;
    //        }

    //        members = _declarations.FindMembers(type);
    //        return members.SingleOrDefault(m => m.IdentifierName == name);
    //    }

    //    public override void EnterVsAssign(VBAParser.VsAssignContext context)
    //    {
    //        /* named parameter syntax */

    //        // one of these is null...
    //        var callStatementA = context.Parent.Parent.Parent as VBAParser.ICS_S_ProcedureOrArrayCallContext;
    //        var callStatementB = context.Parent.Parent.Parent as VBAParser.ICS_S_VariableOrProcedureCallContext;
    //        var callStatementC = context.Parent.Parent.Parent as VBAParser.ICS_B_MemberProcedureCallContext;
    //        var callStatementD = context.Parent.Parent.Parent as VBAParser.ICS_B_ProcedureCallContext;
            
    //        var procedureName = string.Empty;
    //        ParserRuleContext identifierContext = null;
    //        if (callStatementA != null)
    //        {
    //            procedureName = callStatementA.ambiguousIdentifier().GetText();
    //            identifierContext = callStatementA.ambiguousIdentifier();
    //        }
    //        else if(callStatementB != null)
    //        {
    //            procedureName = callStatementB.ambiguousIdentifier().GetText();
    //            identifierContext = callStatementB.ambiguousIdentifier();
    //        }
    //        else if (callStatementC != null)
    //        {
    //            procedureName = callStatementC.ambiguousIdentifier().GetText();
    //            identifierContext = callStatementC.ambiguousIdentifier();
    //        }
    //        else if (callStatementD != null)
    //        {
    //            procedureName = callStatementD.certainIdentifier().GetText();
    //            identifierContext = callStatementD.certainIdentifier();
    //        }

    //        var procedure = FindProcedureDeclaration(procedureName, identifierContext);
    //        if (procedure == null)
    //        {
    //            return;
    //        }

    //        var call = context.implicitCallStmt_InStmt();
    //        var arg = Resolve(call.iCS_S_VariableOrProcedureCall())
    //                  ?? Resolve(call.iCS_S_ProcedureOrArrayCall())
    //                  ?? Resolve(call.iCS_S_DictionaryCall())
    //                  ?? Resolve(call.iCS_S_MembersCall());

    //        if (arg != null)
    //        {
    //            var reference = new IdentifierReference(_qualifiedName, arg.IdentifierName, context.GetSelection(), context, arg);
    //            arg.AddReference(reference);
    //        }
    //    }

    //    private static readonly DeclarationType[] PropertyAccessors =
    //    {
    //        DeclarationType.PropertyGet,
    //        DeclarationType.PropertyLet,
    //        DeclarationType.PropertySet
    //    };

    //    private Declaration FindProcedureDeclaration(string procedureName, ParserRuleContext context, DeclarationType accessor = DeclarationType.PropertyGet)
    //    {
    //        var matches = _declarations[procedureName]
    //            .Where(declaration => ProcedureDeclarations.Contains(declaration.DeclarationType))
    //            .Where(IsInScope)
    //            .ToList();

    //        if (!matches.Any())
    //        {
    //            return null;
    //        }

    //        if (matches.Count == 1)
    //        {
    //            return matches.First();
    //        }

    //        if (matches.All(m => PropertyAccessors.Contains(m.DeclarationType)))
    //        {
    //            return matches.Find(m => m.DeclarationType == accessor);
    //        }

    //        var procedure = GetClosestScopeDeclaration(matches, context);
    //        return procedure;
    //    }

    //    private Declaration FindVariableDeclaration(string procedureName, ParserRuleContext context, DeclarationType accessorType)
    //    {
    //        var matches = _declarations[procedureName]
    //            .Where(declaration => declaration.DeclarationType == DeclarationType.Variable || declaration.DeclarationType == DeclarationType.Parameter)
    //            .Where(IsInScope);

    //        var variable = GetClosestScopeDeclaration(matches, context, accessorType);
    //        return variable;
    //    }

    //    private static readonly DeclarationType[] ProcedureDeclarations = 
    //        {
    //            DeclarationType.Procedure,
    //            DeclarationType.Function,
    //            DeclarationType.PropertyGet,
    //            DeclarationType.PropertyLet,
    //            DeclarationType.PropertySet
    //        };

    //    private bool IsInScope(Declaration declaration)
    //    {
    //        if (declaration.IsBuiltIn && declaration.Accessibility == Accessibility.Global)
    //        {
    //            return true; // global-scope built-in identifiers are always in scope
    //        }

    //        if (declaration.DeclarationType == DeclarationType.Project)
    //        {
    //            return true; // a project name is always in scope anywhere
    //        }

    //        if (declaration.DeclarationType == DeclarationType.Module ||
    //            declaration.DeclarationType == DeclarationType.Class)
    //        {
    //            // todo: access component instancing properties to do this right (class)
    //            // i.e. a private class in another project wouldn't be accessible
    //            return true;
    //        }

    //        if (ProcedureDeclarations.Contains(declaration.DeclarationType))
    //        {
    //            if (declaration.Accessibility == Accessibility.Public 
    //             || declaration.Accessibility == Accessibility.Implicit)
    //            {
    //                var result = _qualifiedName.Project.Equals(declaration.Project);
    //                return result;
    //            }

    //            return declaration.QualifiedName.QualifiedModuleName == _qualifiedName;
    //        }

    //        return declaration.Scope == _currentScope
    //               || declaration.Scope == ModuleScope
    //               || IsGlobalField(declaration) 
    //               || IsGlobalProcedure(declaration);
    //    }

    //    private static readonly Type[] PropertyContexts =
    //    {
    //        typeof (VBAParser.PropertyGetStmtContext),
    //        typeof (VBAParser.PropertyLetStmtContext),
    //        typeof (VBAParser.PropertySetStmtContext)
    //    };

    //    private Declaration GetClosestScopeDeclaration(IEnumerable<Declaration> declarations, ParserRuleContext context, DeclarationType accessorType = DeclarationType.PropertyGet)
    //    {
    //        if (context.Parent.Parent.Parent is VBAParser.AsTypeClauseContext)
    //        {
    //            accessorType = DeclarationType.Class;
    //        }

    //        var matches = declarations as IList<Declaration> ?? declarations.ToList();
    //        if (!matches.Any())
    //        {
    //            return null;
    //        }

    //        // handle indexed property getters
    //        var currentScopeMatches = matches.Where(declaration => declaration.Context != null &&
    //            (declaration.Scope == _currentScope && !PropertyContexts.Contains(declaration.Context.Parent.Parent.GetType()))
    //            || ((declaration.Context != null && declaration.Context.Parent.Parent is VBAParser.PropertyGetStmtContext
    //                && _currentScopeType == DeclarationType.PropertyGet)
    //            || (declaration.Context != null && declaration.Context.Parent.Parent is VBAParser.PropertySetStmtContext
    //                && _currentScopeType == DeclarationType.PropertySet)
    //            || (declaration.Context != null && declaration.Context.Parent.Parent is VBAParser.PropertyLetStmtContext
    //                && _currentScopeType == DeclarationType.PropertyLet)))
    //            .ToList();
    //        if (currentScopeMatches.Count == 1)
    //        {
    //            return currentScopeMatches[0];
    //        }

    //        // note: commented-out because it breaks the UDT member references, but property getters behave strangely still
    //        //var currentScope = matches.SingleOrDefault(declaration =>
    //        //    IsCurrentScopeMember(accessorType, declaration)
    //        //    && (declaration.DeclarationType == accessorType
    //        //        || accessorType == DeclarationType.PropertyGet));

    //        //if (matches.First().IdentifierName == "procedure")
    //        //{
    //        //    // for debugging - "procedure" is both a UDT member and a parameter to a procedure.
    //        //}

    //        if (matches.Count == 1)
    //        {
    //            return matches[0];
    //        }

    //        var moduleScope = matches.SingleOrDefault(declaration => declaration.Scope == ModuleScope);
    //        if (moduleScope != null)
    //        {
    //            return moduleScope;
    //        }

    //        var splitScope = _currentScope.Split('.');
    //        if (splitScope.Length > 2) // Project.Module.Procedure - i.e. if scope is deeper than module-level
    //        {
    //            var scope = splitScope[0] + '.' + splitScope[1];
    //            var scopeMatches = matches.Where(m => m.ParentScope == scope
    //                                                  && (!PropertyAccessors.Contains(m.DeclarationType)
    //                                                      || m.DeclarationType == accessorType)).ToList();
    //            if (scopeMatches.Count == 1)
    //            {
    //                return scopeMatches.Single();
    //            }

    //            // handle standard library member shadowing:
    //            if (!matches.All(m => m.IsBuiltIn))
    //            {
    //                var ambiguousMatches = matches.Where(m => !m.IsBuiltIn
    //                                                          && (!PropertyAccessors.Contains(m.DeclarationType)
    //                                                              || m.DeclarationType == accessorType)).ToList();

    //                if (ambiguousMatches.Count == 1)
    //                {
    //                    return ambiguousMatches.Single();
    //                }
    //            }
    //        }

    //        var memberProcedureCallContext = context.Parent as VBAParser.ICS_B_MemberProcedureCallContext;
    //        if (memberProcedureCallContext != null)
    //        {
    //            return Resolve(memberProcedureCallContext);
    //        }

    //        var implicitCall = context.Parent.Parent as VBAParser.ImplicitCallStmt_InStmtContext;
    //        if (implicitCall != null)
    //        {
    //            return Resolve(implicitCall.iCS_S_VariableOrProcedureCall())
    //                   ?? Resolve(implicitCall.iCS_S_ProcedureOrArrayCall())
    //                   ?? Resolve(implicitCall.iCS_S_DictionaryCall())
    //                   ?? Resolve(implicitCall.iCS_S_MembersCall());
    //        }

    //        return null;
    //    }

    //    private bool IsCurrentScopeMember(DeclarationType accessorType, Declaration declaration)
    //    {
    //        if (declaration.Scope != ModuleScope && accessorType != DeclarationType.Class)
    //        {
    //            return false;
    //        }

    //        switch (accessorType)
    //        {
    //            case DeclarationType.Class:
    //                return declaration.DeclarationType == DeclarationType.Class;

    //            case DeclarationType.PropertySet:
    //                return declaration.DeclarationType != DeclarationType.PropertyGet && declaration.DeclarationType != DeclarationType.PropertyLet;

    //            case DeclarationType.PropertyLet:
    //                return declaration.DeclarationType != DeclarationType.PropertyGet && declaration.DeclarationType != DeclarationType.PropertySet;

    //            case DeclarationType.PropertyGet:
    //                return declaration.DeclarationType != DeclarationType.PropertyLet && declaration.DeclarationType != DeclarationType.PropertySet;

    //            default:
    //                return true;
    //        } 
    //    }

    //    private bool IsGlobalField(Declaration declaration)
    //    {
    //        // a field isn't a field if it's not a variable or a constant.
    //        if (declaration.DeclarationType != DeclarationType.Variable ||
    //            declaration.DeclarationType != DeclarationType.Constant)
    //        {
    //            return false;
    //        }

    //        // a field is only global if it's declared as Public or Global in a standard module.
    //        var moduleMatches = _declarations[declaration.ComponentName].ToList();
    //        var modules = moduleMatches.Where(match => match.DeclarationType == DeclarationType.Module);

    //        // Friend members are only visible within the same project.
    //        var isSameProject = declaration.Project == _qualifiedName.Project;

    //        // todo: verify that this isn't overkill. Friend modifier has restricted legal usage.
    //        return modules.Any()
    //               && (declaration.Accessibility == Accessibility.Global
    //                   || declaration.Accessibility == Accessibility.Public
    //                   || (isSameProject && declaration.Accessibility == Accessibility.Friend));
    //    }

    //    private bool IsGlobalProcedure(Declaration declaration)
    //    {
    //        // a procedure is global if it's a Sub or Function (properties are never global).
    //        // since we have no visibility on module attributes,
    //        // we must assume a class member can be called from a default instance.

    //        if (declaration.DeclarationType != DeclarationType.Procedure ||
    //            declaration.DeclarationType != DeclarationType.Function)
    //        {
    //            return false;
    //        }

    //        // Friend members are only visible within the same project.
    //        var isSameProject = declaration.Project == _qualifiedName.Project;

    //        // implicit (unspecified) accessibility makes a member Public,
    //        // so if it's in the same project, it's public whenever it's not explicitly Private:
    //        return isSameProject && declaration.Accessibility == Accessibility.Friend
    //               || declaration.Accessibility != Accessibility.Private;
    //    }

    //    private bool IsDeclarativeContext(VBAParser.AmbiguousIdentifierContext context)
    //    {
    //        return IsDeclarativeParentContext(context.Parent);
    //    }

    //    private bool IsDeclarativeContext(VBAParser.CertainIdentifierContext context)
    //    {
    //        return IsDeclarativeParentContext(context.Parent);
    //    }

    //    private static readonly Type[] DeclarativeContextTypes =
    //    {
    //        typeof (VBAParser.VariableSubStmtContext),
    //        typeof (VBAParser.ConstSubStmtContext),
    //        typeof (VBAParser.ArgContext),
    //        typeof (VBAParser.SubStmtContext),
    //        typeof (VBAParser.FunctionStmtContext),
    //        typeof (VBAParser.PropertyGetStmtContext),
    //        typeof (VBAParser.PropertyLetStmtContext),
    //        typeof (VBAParser.PropertySetStmtContext),
    //        typeof (VBAParser.TypeStmtContext),
    //        typeof (VBAParser.TypeStmt_ElementContext),
    //        typeof (VBAParser.EnumerationStmtContext),
    //        typeof (VBAParser.EnumerationStmt_ConstantContext),
    //        typeof (VBAParser.DeclareStmtContext),
    //        typeof (VBAParser.EventStmtContext)
    //    };

    //    private bool IsDeclarativeParentContext(RuleContext parentContext)
    //    {
    //        return DeclarativeContextTypes.Contains(parentContext.GetType());
    //    }
    //}
}