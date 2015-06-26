using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReferenceResolver
    {
        private enum ContextAccessorType
        {
            GetValueOrReference,
            AssignValue,
            AssignReference
        }

        private readonly Declarations _declarations;
        private readonly QualifiedModuleName _qualifiedModuleName;

        private readonly IReadOnlyList<DeclarationType> _moduleTypes;
        private readonly IReadOnlyList<DeclarationType> _memberTypes;

        private readonly Stack<Declaration> _withBlockQualifiers;
        private readonly HashSet<RuleContext> _alreadyResolved;

        public IdentifierReferenceResolver(QualifiedModuleName qualifiedModuleName, Declarations declarations)
        {
            _qualifiedModuleName = qualifiedModuleName;
            _declarations = declarations;

            _withBlockQualifiers = new Stack<Declaration>();
            _alreadyResolved = new HashSet<RuleContext>();

            _moduleTypes = new List<DeclarationType>(new[]
            {
                DeclarationType.Module, 
                DeclarationType.Class,
                DeclarationType.Project
            });

            _memberTypes = new List<DeclarationType>(new[]
            {
                DeclarationType.Function, 
                DeclarationType.Procedure, 
                DeclarationType.PropertyGet, 
                DeclarationType.PropertyLet, 
                DeclarationType.PropertySet
            });

        }

        private Declaration _currentScope;

        public void SetCurrentScope()
        {
            _currentScope = _declarations.Items.Single(item =>
                _moduleTypes.Contains(item.DeclarationType)
                && item.Project == _qualifiedModuleName.Project
                && item.ComponentName == _qualifiedModuleName.ComponentName);

            _alreadyResolved.Clear();
        }

        public void SetCurrentScope(string memberName, DeclarationType? accessor = null)
        {
            _currentScope = _declarations.Items.Single(item =>
                _memberTypes.Contains(item.DeclarationType)
                && (!accessor.HasValue || item.DeclarationType == accessor.Value)
                && item.Project == _qualifiedModuleName.Project
                && item.ComponentName == _qualifiedModuleName.ComponentName
                && item.IdentifierName == memberName);
        }

        public void EnterWithBlock(VBAParser.WithStmtContext context)
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
                _alreadyResolved.Add(reference.Context);
            }
            _withBlockQualifiers.Push(qualifier); // note: pushes null if unresolved
        }

        public void ExitWithBlock()
        {
            _withBlockQualifiers.Pop();
        }

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

            var type = ResolveInternal(context.iCS_S_VariableOrProcedureCall(), localScope)
                       ?? ResolveInternal(context.iCS_S_ProcedureOrArrayCall(), localScope)
                       ?? ResolveInternal(context.iCS_S_MembersCall(), localScope)
                       ?? ResolveInternal(dictionaryCall, localScope, ContextAccessorType.GetValueOrReference, dictionaryCall == null ? null : dictionaryCall.dictionaryCallStmt());

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

        private Declaration ResolveInternal(ParserRuleContext callSiteContext, Declaration localScope, ContextAccessorType accessorType = ContextAccessorType.GetValueOrReference, VBAParser.DictionaryCallStmtContext fieldCall = null, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
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
                return ResolveInternal(fieldCall, callee);
            }

            return callee;
        }

        private Declaration ResolveInternal(VBAParser.ICS_S_VariableOrProcedureCallContext context, Declaration localScope, ContextAccessorType accessorType = ContextAccessorType.GetValueOrReference, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (context == null)
            {
                return null;
            }

            var identifierContext = context.ambiguousIdentifier();
            var fieldCall = context.dictionaryCallStmt();

            return ResolveInternal(identifierContext, localScope, accessorType, fieldCall, hasExplicitLetStatement, isAssignmentTarget);
        }

        private Declaration ResolveInternal(VBAParser.DictionaryCallStmtContext fieldCall, Declaration parent, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
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

        private Declaration ResolveInternal(VBAParser.ICS_S_ProcedureOrArrayCallContext context, Declaration localScope, ContextAccessorType accessorType = ContextAccessorType.GetValueOrReference, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (context == null)
            {
                return null;
            }

            var identifierContext = context.ambiguousIdentifier();
            var fieldCall = context.dictionaryCallStmt();
            // todo: understand WTF [baseType] is doing in that grammar rule...

            return ResolveInternal(identifierContext, localScope, accessorType, fieldCall, hasExplicitLetStatement, isAssignmentTarget);
        }

        private Declaration ResolveInternal(VBAParser.ICS_S_MembersCallContext context, ContextAccessorType accessorType, Declaration localScope = null, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (context == null)
            {
                return null;
            }

            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var parent = ResolveInternal(context.iCS_S_ProcedureOrArrayCall(), localScope, accessorType, hasExplicitLetStatement, isAssignmentTarget)
                         ?? ResolveInternal(context.iCS_S_VariableOrProcedureCall(), localScope, accessorType, hasExplicitLetStatement, isAssignmentTarget);

            if (parent != null)
            {
                var parentReference = CreateReference(parent.Context, parent);
                parent.AddReference(parentReference);
                _alreadyResolved.Add(parent.Context);
            }

            var chainedCalls = context.iCS_S_MemberCall();
            foreach (var memberCall in chainedCalls)
            {
                var member = ResolveInternal(memberCall.iCS_S_ProcedureOrArrayCall(), parent, accessorType, hasExplicitLetStatement, isAssignmentTarget)
                             ?? ResolveInternal(memberCall.iCS_S_VariableOrProcedureCall(), parent, accessorType, hasExplicitLetStatement, isAssignmentTarget);

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

            return ResolveInternal(fieldCall, parent, hasExplicitLetStatement, isAssignmentTarget);
        }

        private Declaration ResolveInternal(VBAParser.ImplicitCallStmt_InStmtContext callSiteContext, Declaration localScope, ContextAccessorType accessorType, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (callSiteContext == null)
            {
                return null;
            }

            var dictionaryCall = callSiteContext.iCS_S_DictionaryCall();
            var fieldCall = dictionaryCall == null ? null : dictionaryCall.dictionaryCallStmt();

            return ResolveInternal(callSiteContext.iCS_S_VariableOrProcedureCall(), localScope, accessorType, hasExplicitLetStatement, isAssignmentTarget)
                   ?? ResolveInternal(callSiteContext.iCS_S_ProcedureOrArrayCall(), localScope, accessorType, hasExplicitLetStatement, isAssignmentTarget)
                   ?? ResolveInternal(callSiteContext.iCS_S_MembersCall(), accessorType, localScope, hasExplicitLetStatement, isAssignmentTarget)
                   ?? ResolveInternal(callSiteContext.iCS_S_DictionaryCall(), localScope, accessorType, fieldCall, hasExplicitLetStatement, isAssignmentTarget);
        }

        private Declaration ResolveInternal(VBAParser.ICS_B_ProcedureCallContext context)
        {
            if (context == null)
            {
                return null;
            }

            var identifierContext = context.certainIdentifier();
            var callee = ResolveInternal(identifierContext, _currentScope);
            if (callee == null)
            {
                return null;
            }

            var reference = CreateReference(identifierContext, callee);
            callee.AddReference(reference);
            _alreadyResolved.Add(context);

            return callee;
        }

        private void ResolveIdentifier(VBAParser.AmbiguousIdentifierContext context)
        {
            if (context == null)
            {
                return;
            }

            var identifier = ResolveInternal(context, _currentScope);
            if (identifier == null)
            {
                return;
            }

            identifier.AddReference(CreateReference(context, identifier));
        }

        public void Resolve(VBAParser.ICS_B_ProcedureCallContext context)
        {
            if (_alreadyResolved.Contains(context))
            {
                return;
            }

            ResolveInternal(context);
        }

        public void Resolve(VBAParser.ICS_B_MemberProcedureCallContext context)
        {
            if (_alreadyResolved.Contains(context))
            {
                return;
            }

            var parentScope = ResolveInternal(context.implicitCallStmt_InStmt(), _currentScope, ContextAccessorType.GetValueOrReference);
            var parentType = ResolveType(parentScope);

            if (_withBlockQualifiers.Any())
            {
                parentType = _withBlockQualifiers.Peek();
                parentScope = ResolveInternal(context.implicitCallStmt_InStmt(), parentType, ContextAccessorType.GetValueOrReference)
                              ?? ResolveInternal(context.ambiguousIdentifier(), parentType);
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
            ResolveInternal(fieldCall, member);
        }

        public void Resolve(VBAParser.ICS_S_VariableOrProcedureCallContext context)
        {
            TryResolve(context);
        }

        public void Resolve(VBAParser.ICS_S_ProcedureOrArrayCallContext context)
        {
            TryResolve(context);
        }

        public void Resolve(VBAParser.ICS_S_MembersCallContext context)
        {
            TryResolve(context);
        }

        public void Resolve(VBAParser.ICS_S_DictionaryCallContext context)
        {
            TryResolve(context);
        }

        private void TryResolve<TContext>(TContext context) where TContext : ParserRuleContext
        {
            if (context == null || _alreadyResolved.Contains(context))
            {
                return;
            }

            try
            {
                ResolveInternal(context, _currentScope);
            }
            catch (InvalidOperationException)
            {
                // bug: more than a single match was found.
            }
        }

        public void Resolve(VBAParser.LetStmtContext context)
        {
            var leftSide = context.implicitCallStmt_InStmt();
            var letStatement = context.LET();
            ResolveInternal(leftSide, _currentScope, ContextAccessorType.AssignValue, letStatement != null, true);
        }

        public void Resolve(VBAParser.SetStmtContext context)
        {
            var leftSide = context.implicitCallStmt_InStmt();
            ResolveInternal(leftSide, _currentScope, ContextAccessorType.AssignReference, false, true);
        }

        public void Resolve(VBAParser.AsTypeClauseContext context)
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

        public void Resolve(VBAParser.ForNextStmtContext context)
        {
            var identifiers = context.ambiguousIdentifier();
            var identifier = ResolveInternal(identifiers[0], _currentScope, ContextAccessorType.AssignValue, null, false, true);
            
            // each iteration counts as an assignment
            var reference = CreateReference(identifiers[0], identifier, true);
            identifier.AddReference(reference);

            if (identifiers.Count > 1)
            {
                identifier.AddReference(CreateReference(identifiers[1], identifier));
            }
        }

        public void Resolve(VBAParser.ForEachStmtContext context)
        {
            var identifiers = context.ambiguousIdentifier();
            var identifier = ResolveInternal(identifiers[0], _currentScope, ContextAccessorType.AssignValue, null, false, true);

            // each iteration counts as an assignment
            var reference = CreateReference(identifiers[0], identifier, true);
            identifier.AddReference(reference);

            if (identifiers.Count > 1)
            {
                identifier.AddReference(CreateReference(identifiers[1], identifier));
            }
        }

        public void Resolve(VBAParser.ImplementsStmtContext context)
        {
            ResolveIdentifier(context.ambiguousIdentifier());
        }

        public void Resolve(VBAParser.RaiseEventStmtContext context)
        {
            ResolveIdentifier(context.ambiguousIdentifier());
        }

        public void Resolve(VBAParser.ResumeStmtContext context)
        {
            ResolveIdentifier(context.ambiguousIdentifier());
        }

        public void Resolve(VBAParser.FileNumberContext context)
        {
            ResolveIdentifier(context.ambiguousIdentifier());
        }

        public void Resolve(VBAParser.ArgDefaultValueContext context)
        {
            ResolveIdentifier(context.ambiguousIdentifier());
        }

        public void Resolve(VBAParser.FieldLengthContext context)
        {
            ResolveIdentifier(context.ambiguousIdentifier());
        }

        public void Resolve(VBAParser.VsAssignContext context)
        {
            // named parameter reference must be scoped to called procedure
            var callee = FindParentCall(context);
            ResolveInternal(context.implicitCallStmt_InStmt(), callee);
        }

        private Declaration FindParentCall(VBAParser.VsAssignContext context)
        {
            var calleeContext = context.Parent.Parent.Parent;
            return ResolveInternal(calleeContext as VBAParser.ICS_B_ProcedureCallContext)
                   ?? ResolveInternal(calleeContext as VBAParser.ICS_S_VariableOrProcedureCallContext, _currentScope)
                   ?? ResolveInternal(calleeContext as VBAParser.ICS_S_ProcedureOrArrayCallContext, _currentScope)
                   ?? ResolveInternal(calleeContext as VBAParser.ICS_S_MembersCallContext, _currentScope);
        }

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

        private Declaration FindModuleScopeProcedure(string identifierName, Declaration localScope, ContextAccessorType accessorType, bool isAssignmentTarget = false)
        {
            if (localScope == null)
            {
                localScope = _currentScope;
            }

            return _declarations[identifierName].SingleOrDefault(item =>
                IsProcedure(item) || IsPropertyAccessor(item, accessorType, localScope, isAssignmentTarget));
        }

        private Declaration FindProjectScopeDeclaration(string identifierName)
        {
            // assume unqualified variable call, i.e. unique declaration.
            return _declarations[identifierName].SingleOrDefault(item =>
                (item.Accessibility == Accessibility.Public
                 || item.Accessibility == Accessibility.Global
                 || _moduleTypes.Contains(item.DeclarationType)));
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
            if (item.Equals(localScope))
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
    }
}