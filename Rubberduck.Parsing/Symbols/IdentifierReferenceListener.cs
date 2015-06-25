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
        private readonly HashSet<RuleContext> _alreadyResolved;

        public IdentifierReferenceListener(QualifiedModuleName qualifiedModuleName, Declarations declarations)
        {
            _qualifiedModuleName = qualifiedModuleName;
            _declarations = declarations;

            _moduleTypes = new HashSet<DeclarationType>(new[]
            {
                DeclarationType.Module, 
                DeclarationType.Class,
                DeclarationType.Project
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
            _alreadyResolved = new HashSet<RuleContext>();
            
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

            _alreadyResolved.Clear();
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
            Declaration qualifier = null;
            IdentifierReference reference = null;

            if (context.NEW() == null)
            {
                // with block is using an identifier declared elsewhere.
                qualifier = ResolveType(context.implicitCallStmt_InStmt());
                reference = CreateReference(context.implicitCallStmt_InStmt(), qualifier);
            }
            else
            {
                // with block is using an anonymous declaration.
                // i.e. object variable reference is held by the with block itself.
                var typeContext = context.type();
                var baseTypeContext = typeContext.baseType();
                if (baseTypeContext != null)
                {
                    var collectionContext = baseTypeContext.COLLECTION();
                    if (collectionContext != null)
                    {
                        // object variable is a built-in Collection class instance
                        qualifier = _declarations.Items.Single(item => item.IsBuiltIn
                            && item.IdentifierName == collectionContext.GetText()
                            && item.DeclarationType == DeclarationType.Class);
                        reference = CreateReference(baseTypeContext, qualifier);
                    }
                }
                else
                {
                    qualifier = ResolveType(typeContext.complexType());
                    reference = CreateReference(typeContext.complexType(), qualifier);
                }
            }

            if (qualifier != null)
            {
                qualifier.AddReference(reference);
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
                    && _moduleTypes.Contains(item.DeclarationType));
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

            var dictionaryCall = context.iCS_S_DictionaryCall();

            var type = Resolve(context.iCS_S_VariableOrProcedureCall(), localScope)
                ?? Resolve(context.iCS_S_ProcedureOrArrayCall(), localScope)
                ?? Resolve(context.iCS_S_MembersCall(), localScope)
                ?? Resolve(dictionaryCall, localScope, ContextAccessorType.GetValueOrReference, dictionaryCall == null ? null : dictionaryCall.dictionaryCallStmt());

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
                && item.ProjectName == _currentScope.ProjectName
                && _moduleTypes.Contains(item.DeclarationType));

            if (result == null)
            {
                // look in all projects (including VbaStdLib library).
                result = _declarations[parent.AsTypeName].SingleOrDefault(item =>
                    _moduleTypes.Contains(item.DeclarationType)
                    && _moduleTypes.Contains(item.DeclarationType));
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
                         ?? FindModuleScopeProcedure(identifierName, localScope, accessorType, isAssignmentTarget)
                         ?? FindProjectScopeDeclaration(identifierName);

            if (callee == null)
            {
                // calls inside With block can still refer to identifiers in _currentScope
                localScope = _currentScope;
                identifierName = callSiteContext.GetText();
                callee = FindLocalScopeDeclaration(identifierName, localScope)
                         ?? FindModuleScopeDeclaration(identifierName, localScope)
                         ?? FindModuleScopeProcedure(identifierName, localScope, accessorType, isAssignmentTarget)
                         ?? FindProjectScopeDeclaration(identifierName);
            }

            if (callee == null)
            {
                return null;
            }

            var reference = CreateReference(callSiteContext, callee, isAssignmentTarget, hasExplicitLetStatement);
            callee.AddReference(reference);
            _alreadyResolved.Add(callSiteContext.Parent);

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

            var identifierContext = context.ambiguousIdentifier();
            var fieldCall = context.dictionaryCallStmt();

            return Resolve(identifierContext, localScope, accessorType, fieldCall, hasExplicitLetStatement, isAssignmentTarget);
        }

        private Declaration Resolve(VBAParser.DictionaryCallStmtContext fieldCall, Declaration parent, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (fieldCall == null)
            {
                return null;
            }

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
            _alreadyResolved.Add(fieldCall);

            return result;
        }

        private Declaration Resolve(VBAParser.ICS_S_ProcedureOrArrayCallContext context, Declaration localScope, ContextAccessorType accessorType = ContextAccessorType.GetValueOrReference, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (context == null)
            {
                return null;
            }

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

            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var parent = Resolve(context.iCS_S_ProcedureOrArrayCall(), localScope, accessorType, hasExplicitLetStatement, isAssignmentTarget)
                      ?? Resolve(context.iCS_S_VariableOrProcedureCall(), localScope, accessorType, hasExplicitLetStatement, isAssignmentTarget);

            if (parent != null)
            {
                var parentReference = CreateReference(parent.Context, parent);
                parent.AddReference(parentReference);
                _alreadyResolved.Add(parent.Context);
            }

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

        private Declaration Resolve(VBAParser.ICS_B_ProcedureCallContext context)
        {
            if (context == null)
            {
                return null;
            }

            var identifierContext = context.certainIdentifier();
            var callee = Resolve(identifierContext, _currentScope);
            if (callee == null)
            {
                return null;
            }

            var reference = CreateReference(identifierContext, callee);
            callee.AddReference(reference);
            _alreadyResolved.Add(context);

            return callee;
        }

        private Declaration Resolve(VBAParser.ImplicitCallStmt_InStmtContext callSiteContext, Declaration localScope, ContextAccessorType accessorType, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (callSiteContext == null)
            {
                return null;
            }

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
                item.ParentScope == localScope.Scope
                && !_moduleTypes.Contains(item.DeclarationType));
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
                && !item.DeclarationType.HasFlag(DeclarationType.Member)
                && !_moduleTypes.Contains(item.DeclarationType));
        }

        /// <summary>
        /// Finds a procedure declaration in the specified scope.
        /// </summary>
        /// <param name="identifierName">The name of the identifier to find.</param>
        /// <param name="localScope">The scope considered local.</param>
        /// <param name="accessorType">Disambiguates <see cref="DeclarationType.PropertyLet"/> and <see cref="DeclarationType.PropertySet"/> accessors.</param>
        private Declaration FindModuleScopeProcedure(string identifierName, Declaration localScope, ContextAccessorType accessorType, bool isAssignmentTarget = false)
        {
            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var result = _declarations[identifierName].Where(item =>
                IsProcedure(item) || IsPropertyAccessor(item, accessorType, localScope, isAssignmentTarget))
                  .ToList();

            return result.SingleOrDefault();
        }

        private bool IsProcedure(Declaration item)
        {
            return item.DeclarationType == DeclarationType.Procedure
                   || item.DeclarationType == DeclarationType.Function;
        }

        private bool IsPropertyAccessor(Declaration item, ContextAccessorType accessorType, Declaration localScope, bool isAssignmentTarget = false)
        {
            var isProperty = item.DeclarationType.HasFlag(DeclarationType.Property);
            if (!isProperty)
            {
                return false;
            }

            if (item.Equals(localScope) && item.DeclarationType == DeclarationType.PropertyGet)
            {
                // we're resolving the getter's return value assignment
                return true;
            }
            if(item.Equals(localScope))
            {
                // getter can't reference setter.. right?
                return false;
            }

            return (accessorType == ContextAccessorType.AssignValue &&
                    item.DeclarationType == DeclarationType.PropertyLet)
                   ||
                   (accessorType == ContextAccessorType.AssignReference &&
                    item.DeclarationType == DeclarationType.PropertySet)
                   ||
                   (accessorType == ContextAccessorType.GetValueOrReference &&
                    item.DeclarationType == DeclarationType.PropertyGet &&
                    !isAssignmentTarget);
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
                (item.Accessibility == Accessibility.Public
                || item.Accessibility == Accessibility.Global
                || item.DeclarationType == DeclarationType.Module));
        }

        #endregion

        #region IVBAListener overrides

        public override void EnterICS_B_ProcedureCall(VBAParser.ICS_B_ProcedureCallContext context)
        {
            Resolve(context.certainIdentifier(), _currentScope);
        }

        public override void EnterICS_B_MemberProcedureCall(VBAParser.ICS_B_MemberProcedureCallContext context)
        {
            if (_alreadyResolved.Contains(context))
            {
                return;
            }

            var parentScope = Resolve(context.implicitCallStmt_InStmt(), _currentScope, ContextAccessorType.GetValueOrReference);
            var parentType = ResolveType(parentScope);

            if (_withBlockQualifiers.Any())
            {
                parentType = _withBlockQualifiers.Peek();
                parentScope = Resolve(context.implicitCallStmt_InStmt(), parentType, ContextAccessorType.GetValueOrReference);
                parentType = ResolveType(parentScope);
            }
            if (parentType == null)
            {
                return;
            }

            var identifierContext = context.ambiguousIdentifier();
            var member = _declarations[identifierContext.GetText()].SingleOrDefault(item =>
                item.ComponentName == parentType.ComponentName);

            if (member != null)
            {
                var reference = CreateReference(identifierContext, member);
                member.AddReference(reference);
                _alreadyResolved.Add(context);
            }
            
            var fieldCall = context.dictionaryCallStmt();
            Resolve(fieldCall, member);
        }

        public override void EnterICS_S_VariableOrProcedureCall(VBAParser.ICS_S_VariableOrProcedureCallContext context)
        {
            if (_alreadyResolved.Contains(context))
            {
                return;
            }

            try
            {
                Resolve(context, _currentScope);
            }
            catch (InvalidOperationException)
            {
                // more than a single match was found.
            }
        }

        public override void EnterICS_S_ProcedureOrArrayCall(VBAParser.ICS_S_ProcedureOrArrayCallContext context)
        {
            if (_alreadyResolved.Contains(context))
            {
                return;
            }

            try
            {
                Resolve(context, _currentScope);
            }
            catch (InvalidOperationException)
            {
                // more than a single match was found.
            }
        }

        public override void EnterICS_S_MembersCall(VBAParser.ICS_S_MembersCallContext context)
        {
            if (_alreadyResolved.Contains(context))
            {
                return;
            }

            try
            {
                Resolve(context, _currentScope);
            }
            catch (InvalidOperationException)
            {
                // more than a single match was found.
            }
        }

        public override void EnterICS_S_DictionaryCall(VBAParser.ICS_S_DictionaryCallContext context)
        {
            if (_alreadyResolved.Contains(context))
            {
                return;
            }

            try
            {
                Resolve(context, _currentScope);
            }
            catch (InvalidOperationException)
            {
                // more than a single match was found.
            }
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

        public override void EnterAsTypeClause(VBAParser.AsTypeClauseContext context)
        {
            var asType = context.type();
            if (asType == null)
            {
                return;
            }

            Declaration type = null;
            IdentifierReference reference = null;

            var baseType = asType.baseType();
            if (baseType != null)
            {
                var collection = baseType.COLLECTION();
                if (collection != null)
                {
                    type = _declarations[collection.GetText()].SingleOrDefault(item => item.IsBuiltIn && item.DeclarationType == DeclarationType.Class);
                    reference = CreateReference(baseType, type);
                }
            }
            else
            {
                type = ResolveType(asType.complexType());
                reference = CreateReference(asType.complexType(), type);
            }

            if (type != null)
            {
                type.AddReference(reference);
            }
        }

        public override void EnterForNextStmt(VBAParser.ForNextStmtContext context)
        {
            var identifiers = context.ambiguousIdentifier();
            var identifier = Resolve(identifiers[0], _currentScope, ContextAccessorType.AssignValue, null, false, true);
            
            // each iteration counts as an assignment
            var reference = CreateReference(identifiers[0], identifier, true);
            identifier.AddReference(reference);

            if (identifiers.Count > 1)
            {
                identifier.AddReference(CreateReference(identifiers[1], identifier));
            }
        }

        public override void EnterForEachStmt(VBAParser.ForEachStmtContext context)
        {
            var identifiers = context.ambiguousIdentifier();
            var identifier = Resolve(identifiers[0], _currentScope, ContextAccessorType.AssignValue, null, false, true);

            // each iteration counts as an assignment
            var reference = CreateReference(identifiers[0], identifier, true);
            identifier.AddReference(reference);

            if (identifiers.Count > 1)
            {
                identifier.AddReference(CreateReference(identifiers[1], identifier));
            }
        }

        public override void EnterImplementsStmt(VBAParser.ImplementsStmtContext context)
        {
            var identifierContext = context.ambiguousIdentifier();
            var identifier = Resolve(identifierContext, _currentScope);
            identifier.AddReference(CreateReference(identifierContext, identifier));
        }

        public override void EnterRaiseEventStmt(VBAParser.RaiseEventStmtContext context)
        {
            var identifierContext = context.ambiguousIdentifier();
            var identifier = Resolve(identifierContext, _currentScope);
            identifier.AddReference(CreateReference(identifierContext, identifier));
        }

        public override void EnterResumeStmt(VBAParser.ResumeStmtContext context)
        {
            var identifierContext = context.ambiguousIdentifier();
            var identifier = Resolve(identifierContext, _currentScope);
            identifier.AddReference(CreateReference(identifierContext, identifier));
        }

        public override void EnterFileNumber(VBAParser.FileNumberContext context)
        {
            var identifierContext = context.ambiguousIdentifier();
            if (identifierContext == null)
            {
                return;
            }
            var identifier = Resolve(identifierContext, _currentScope);
            identifier.AddReference(CreateReference(identifierContext, identifier));
        }

        public override void EnterArgDefaultValue(VBAParser.ArgDefaultValueContext context)
        {
            var identifierContext = context.ambiguousIdentifier();
            if (identifierContext == null)
            {
                return;
            }
            var identifier = Resolve(identifierContext, _currentScope);
            identifier.AddReference(CreateReference(identifierContext, identifier));
        }

        public override void EnterFieldLength(VBAParser.FieldLengthContext context)
        {
            var identifierContext = context.ambiguousIdentifier();
            if (identifierContext == null)
            {
                return;
            }
            var identifier = Resolve(identifierContext, _currentScope);
            identifier.AddReference(CreateReference(identifierContext, identifier));
        }

        public override void EnterVsAssign(VBAParser.VsAssignContext context)
        {
            // named parameter reference must be scoped to called procedure
            var callee = FindParentCall(context);
            Resolve(context.implicitCallStmt_InStmt(), callee);
        }

        private Declaration FindParentCall(VBAParser.VsAssignContext context)
        {
            var calleeContext = context.Parent.Parent.Parent;
            return Resolve(calleeContext as VBAParser.ICS_B_ProcedureCallContext)
                   ?? Resolve(calleeContext as VBAParser.ICS_S_VariableOrProcedureCallContext, _currentScope)
                   ?? Resolve(calleeContext as VBAParser.ICS_S_ProcedureOrArrayCallContext, _currentScope)
                   ?? Resolve(calleeContext as VBAParser.ICS_S_MembersCallContext, _currentScope);
        }

        #endregion

        private IdentifierReference CreateReference(ParserRuleContext callSiteContext, Declaration callee, bool isAssignmentTarget = false, bool hasExplicitLetStatement = false)
        {
            if (callSiteContext == null)
            {
                return null;
            }
            var name = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            return new IdentifierReference(_qualifiedModuleName, name, selection, callSiteContext, callee, isAssignmentTarget, hasExplicitLetStatement);
        }
    }
}