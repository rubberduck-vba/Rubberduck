using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
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

        private readonly IEnumerable<Declaration> _declarations;
        private readonly IEnumerable<CommentNode> _comments;

        private readonly QualifiedModuleName _qualifiedModuleName;

        private readonly IReadOnlyList<DeclarationType> _moduleTypes;
        private readonly IReadOnlyList<DeclarationType> _memberTypes;
        private readonly IReadOnlyList<DeclarationType> _returningMemberTypes;

        private readonly IReadOnlyList<Accessibility> _projectScopePublicModifiers; 

        private readonly Stack<Declaration> _withBlockQualifiers;
        private readonly HashSet<RuleContext> _alreadyResolved;

        public IdentifierReferenceResolver(QualifiedModuleName qualifiedModuleName, IEnumerable<Declaration> declarations, IEnumerable<CommentNode> comments)
        {
            _qualifiedModuleName = qualifiedModuleName;
            _declarations = declarations;
            _comments = comments;

            _withBlockQualifiers = new Stack<Declaration>();
            _alreadyResolved = new HashSet<RuleContext>();

            _moduleTypes = new List<DeclarationType>(new[]
            {
                DeclarationType.Module, 
                DeclarationType.Class,
            });

            _memberTypes = new List<DeclarationType>(new[]
            {
                DeclarationType.Function, 
                DeclarationType.Procedure, 
                DeclarationType.PropertyGet, 
                DeclarationType.PropertyLet, 
                DeclarationType.PropertySet,
            });

            _returningMemberTypes = new List<DeclarationType>(new[]
            {
                DeclarationType.Function,
                DeclarationType.PropertyGet, 
            });

            _projectScopePublicModifiers = new List<Accessibility>(new[]
            {
                Accessibility.Public, 
                Accessibility.Global, 
                Accessibility.Friend, 
                Accessibility.Implicit, 
            });

            SetCurrentScope();
        }

        private Declaration _currentScope;

        public void SetCurrentScope()
        {
            _currentScope = _declarations.SingleOrDefault(item => 
                _moduleTypes.Contains(item.DeclarationType)
                && item.Project == _qualifiedModuleName.Project
                && item.ComponentName == _qualifiedModuleName.ComponentName);

            _alreadyResolved.Clear();
        }

        public void SetCurrentScope(string memberName, DeclarationType? accessor = null)
        {
            _currentScope = _declarations.SingleOrDefault(item =>
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
                qualifier = ResolveInternal(context.implicitCallStmt_InStmt(), _currentScope, ContextAccessorType.GetValueOrReference);
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
                        qualifier = _declarations.Single(item => item.IsBuiltIn
                                                                       && item.IdentifierName == collectionContext.GetText()
                                                                       && item.DeclarationType == DeclarationType.Class);
                        reference = CreateReference(baseTypeContext, qualifier);
                    }
                }
                else
                {
                    //qualifier = ResolveType(typeContext.complexType());
                }
            }

            if (qualifier != null && reference != null)
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
            if (callSiteContext == null || _currentScope == null)
            {
                return null;
            }
            var name = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var annotations = FindAnnotations(selection.StartLine);
            return new IdentifierReference(_qualifiedModuleName, _currentScope.Scope, name, selection, callSiteContext, callee, isAssignmentTarget, hasExplicitLetStatement, annotations);
        }

        private string FindAnnotations(int line)
        {
            if (_comments == null)
            {
                return null;
            }

            var commentAbove = _comments.SingleOrDefault(comment => comment.QualifiedSelection.Selection.EndLine == line - 1);
            if (commentAbove != null && commentAbove.CommentText.StartsWith("@"))
            {
                return commentAbove.CommentText;
            }
            return null;
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
                var matches = _declarations.Where(d => d.IdentifierName == identifier.GetText());
                try
                {
                    return matches.SingleOrDefault(item =>
                        item.ProjectName == libraryName
                        && _projectScopePublicModifiers.Contains(item.Accessibility)
                        && _moduleTypes.Contains(item.DeclarationType)
                        || (_currentScope != null && _memberTypes.Contains(_currentScope.DeclarationType) 
                            && item.DeclarationType == DeclarationType.UserDefinedType
                            && item.ComponentName == _currentScope.ComponentName));
                }
                catch (InvalidOperationException)
                {
                    return null;
                }
            }

            return null;
        }

        private Declaration ResolveType(Declaration parent)
        {
            if (parent == null || parent.AsTypeName == null)
            {
                return null;
            }

            var identifier = parent.AsTypeName.Contains(".")
                ? parent.AsTypeName.Split('.').Last()
                : parent.AsTypeName;

            var matches = _declarations.Where(d => d.IdentifierName == identifier).ToList();

            try
            {
                var result = matches.SingleOrDefault(item =>
                    item.DeclarationType == DeclarationType.UserDefinedType
                    && item.Project == _currentScope.Project
                    && item.ComponentName == _currentScope.ComponentName);

                if (result == null)
                {
                    result = matches.SingleOrDefault(item =>
                        _moduleTypes.Contains(item.DeclarationType)
                        && item.Project == _currentScope.Project);                
                }

                if (result == null)
                {
                    result = matches.SingleOrDefault(item =>
                        _moduleTypes.Contains(item.DeclarationType));
                }

                return result;
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        private static readonly Type[] IdentifierContexts =
        {
            typeof (VBAParser.AmbiguousIdentifierContext),
            typeof (VBAParser.CertainIdentifierContext)
        };

        private Declaration ResolveInternal(ParserRuleContext callSiteContext, Declaration localScope, ContextAccessorType accessorType = ContextAccessorType.GetValueOrReference, VBAParser.DictionaryCallStmtContext fieldCall = null, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (callSiteContext == null || _alreadyResolved.Contains(callSiteContext))
            {
                return null;
            }

            if (!IdentifierContexts.Contains(callSiteContext.GetType()))
            {
                throw new ArgumentException("'" + callSiteContext.GetType().Name + "' is not an identifier context.", "callSiteContext");
            }

            if (localScope == null)
            {
                localScope = _currentScope;
            }

            if (localScope == null)
            {
                return null;
            }

            var parentContext = callSiteContext.Parent;
            var identifierName = callSiteContext.GetText();

            var sibling = parentContext.ChildCount > 1 ? parentContext.GetChild(1) : null;
            var hasStringQualifier = sibling is VBAParser.TypeHintContext && sibling.GetText() == "$";

            Declaration callee = null;
            if (localScope.DeclarationType == DeclarationType.Variable)
            {
                // localScope is probably a UDT
                var udt = ResolveType(localScope);
                if (udt != null && udt.DeclarationType == DeclarationType.UserDefinedType)
                {
                    callee = _declarations.Where(d => d.IdentifierName == identifierName).SingleOrDefault(item => item.Context != null && item.Context.Parent == udt.Context);
                }
            }
            else
            {
                callee = Resolve(identifierName, localScope, accessorType, parentContext is VBAParser.ICS_S_VariableOrProcedureCallContext, isAssignmentTarget, hasStringQualifier);
            }

            if (callee == null)
            {
                // calls inside With block can still refer to identifiers in _currentScope
                localScope = _currentScope;
                identifierName = callSiteContext.GetText();
                callee = FindLocalScopeDeclaration(identifierName, localScope, parentContext is VBAParser.ICS_S_VariableOrProcedureCallContext, isAssignmentTarget)
                      ?? FindModuleScopeProcedure(identifierName, localScope, accessorType, isAssignmentTarget)
                      ?? FindModuleScopeDeclaration(identifierName, localScope)
                      ?? FindProjectScopeDeclaration(identifierName, Equals(localScope, _currentScope) ? null : localScope, hasStringQualifier);
            }

            if (callee == null)
            {
                return null;
            }

            var reference = CreateReference(callSiteContext, callee, isAssignmentTarget, hasExplicitLetStatement);
            callee.AddReference(reference);
            _alreadyResolved.Add(reference.Context);
            _alreadyResolved.Add(callSiteContext);

            if (fieldCall != null)
            {
                return ResolveInternal(fieldCall, callee);
            }

            return callee;
        }

        private Declaration Resolve(string identifierName, Declaration localScope, ContextAccessorType accessorType, bool parentContextIsVariableOrProcedureCall = false, bool isAssignmentTarget = false, bool hasStringQualifier = false)
        {
            return FindLocalScopeDeclaration(identifierName, localScope, parentContextIsVariableOrProcedureCall, isAssignmentTarget)
                ?? FindModuleScopeProcedure(identifierName, localScope, accessorType, isAssignmentTarget)
                ?? FindModuleScopeDeclaration(identifierName, localScope)
                ?? FindProjectScopeDeclaration(identifierName, Equals(localScope, _currentScope) ? null : localScope, hasStringQualifier);
        }

        private Declaration ResolveInternal(VBAParser.ICS_S_VariableOrProcedureCallContext context, Declaration localScope, ContextAccessorType accessorType = ContextAccessorType.GetValueOrReference, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (context == null)
            {
                return null;
            }

            var identifierContext = context.ambiguousIdentifier();
            var fieldCall = context.dictionaryCallStmt();

            var result = ResolveInternal(identifierContext, localScope, accessorType, fieldCall, hasExplicitLetStatement, isAssignmentTarget);
            if (result != null && localScope != null && !localScope.DeclarationType.HasFlag(DeclarationType.Member))
            {
                localScope.AddMemberCall(CreateReference(context.ambiguousIdentifier(), result));
            }

            return result;
        }

        private Declaration ResolveInternal(VBAParser.DictionaryCallStmtContext fieldCall, Declaration parent, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (fieldCall == null)
            {
                return null;
            }

            var parentType = ResolveType(parent);
            if (parentType == null)
            {
                return null;
            }

            var members = _declarations.Where(declaration => declaration.ParentScope == parentType.Scope);
            var fieldName = fieldCall.ambiguousIdentifier().GetText();

            var result = members.SingleOrDefault(member => member.IdentifierName == fieldName);
            if (result == null)
            {
                return null;
            }

            var identifierContext = fieldCall.ambiguousIdentifier();
            var reference = CreateReference(identifierContext, result, isAssignmentTarget, hasExplicitLetStatement);
            result.AddReference(reference);
            _alreadyResolved.Add(reference.Context);

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

            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var result = ResolveInternal(identifierContext, localScope, accessorType, fieldCall, hasExplicitLetStatement, isAssignmentTarget);
            if (result != null && !localScope.DeclarationType.HasFlag(DeclarationType.Member))
            {
                localScope.AddMemberCall(CreateReference(context.ambiguousIdentifier(), result));
            }

            return result;
        }

        private Declaration ResolveInternal(VBAParser.ICS_S_MembersCallContext context, ContextAccessorType accessorType, Declaration localScope = null, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (context == null)
            {
                return null;
            }

            Declaration parent;
            if (_withBlockQualifiers.Any())
            {
                parent = _withBlockQualifiers.Peek();
            }
            else
            {
                if (localScope == null)
                {
                    localScope = _currentScope;
                }
                parent = ResolveInternal(context.iCS_S_ProcedureOrArrayCall(), localScope, accessorType, hasExplicitLetStatement)
                      ?? ResolveInternal(context.iCS_S_VariableOrProcedureCall(), localScope, accessorType, hasExplicitLetStatement);
            }

            var chainedCalls = context.iCS_S_MemberCall();
            var lastCall = chainedCalls.Last();
            foreach (var memberCall in chainedCalls)
            {
                // if we're on the left side of an assignment, only the last memberCall is the assignment target.
                var isLast = memberCall.Equals(lastCall);
                var accessor = isLast
                    ? accessorType 
                    : ContextAccessorType.GetValueOrReference;
                var isTarget = isLast && isAssignmentTarget;

                var parentType = ResolveType(parent);
                var member = ResolveInternal(memberCall.iCS_S_ProcedureOrArrayCall(), parentType, accessor, hasExplicitLetStatement, isTarget)
                             ?? ResolveInternal(memberCall.iCS_S_VariableOrProcedureCall(), parentType, accessor, hasExplicitLetStatement, isTarget);

                if (member == null)
                {
                    return null;
                }

                member.AddMemberCall(CreateReference(GetMemberCallIdentifierContext(memberCall), parent));
                parent = ResolveType(member);
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
            _alreadyResolved.Add(reference.Context);

            return callee;
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
                parentType = ResolveType(_withBlockQualifiers.Peek());
                parentScope = ResolveInternal(context.implicitCallStmt_InStmt(), parentType, ContextAccessorType.GetValueOrReference)
                              ?? ResolveInternal(context.ambiguousIdentifier(), parentType);
                parentType = ResolveType(parentScope);
            }
            if (parentType == null)
            {
                return;
            }

            var identifierContext = context.ambiguousIdentifier();
            var member = _declarations.Where(d => d.IdentifierName == identifierContext.GetText())
                .SingleOrDefault(item => item.ComponentName == parentType.ComponentName);

            if (member != null)
            {
                var reference = CreateReference(identifierContext, member);

                parentScope.AddMemberCall(CreateReference(context.ambiguousIdentifier(), member));
                member.AddReference(reference);
                _alreadyResolved.Add(reference.Context);
            }
            
            var fieldCall = context.dictionaryCallStmt();
            ResolveInternal(fieldCall, member);
        }

        public void Resolve(VBAParser.ICS_S_VariableOrProcedureCallContext context)
        {
            ResolveInternal(context, _currentScope);
        }

        public void Resolve(VBAParser.ICS_S_ProcedureOrArrayCallContext context)
        {
            ResolveInternal(context, _currentScope);
        }

        public void Resolve(VBAParser.ICS_S_MembersCallContext context)
        {
            if (context == null || _alreadyResolved.Contains(context))
            {
                return;
            }

            Declaration parent;
            if (_withBlockQualifiers.Any())
            {
                parent = ResolveType(_withBlockQualifiers.Peek());
            }
            else
            {
                parent = ResolveInternal(context.iCS_S_ProcedureOrArrayCall(), _currentScope)
                          ?? ResolveInternal(context.iCS_S_VariableOrProcedureCall(), _currentScope);
                parent = ResolveType(parent);
            }

            if (parent != null && parent.Context != null)
            {
                var identifierContext = ((dynamic)parent.Context).ambiguousIdentifier() as VBAParser.AmbiguousIdentifierContext;

                var parentReference = CreateReference(identifierContext, parent);
                parent.AddReference(parentReference);
                _alreadyResolved.Add(parentReference.Context);
            }

            var chainedCalls = context.iCS_S_MemberCall();
            foreach (var memberCall in chainedCalls)
            {
                var notationToken = memberCall.children[0];
                if (notationToken.GetText() == "!")
                {
                    // the memberCall is a shorthand reference to the type's default member.
                    // since the reference isn't explicit, we don't need to care for it.
                    // (and we couldn't handle it if we wanted to, since we aren't parsing member attributes)
                    return;
                }

                var member = ResolveInternal(memberCall.iCS_S_ProcedureOrArrayCall(), parent)
                          ?? ResolveInternal(memberCall.iCS_S_VariableOrProcedureCall(), parent);

                if (member == null)
                {
                    return;
                }

                member.AddMemberCall(CreateReference(GetMemberCallIdentifierContext(memberCall), member));
                parent = ResolveType(member);
            }

            var fieldCall = context.dictionaryCallStmt();
            if (fieldCall == null)
            {
                return;
            }

            ResolveInternal(fieldCall, parent);
            _alreadyResolved.Add(context);
        }

        private VBAParser.AmbiguousIdentifierContext GetMemberCallIdentifierContext(VBAParser.ICS_S_MemberCallContext callContext)
        {
            if (callContext == null)
            {
                return null;
            }

            var procedureOrArrayCall = callContext.iCS_S_ProcedureOrArrayCall();
            if (procedureOrArrayCall != null)
            {
                return procedureOrArrayCall.ambiguousIdentifier();
            }

            var variableOrProcedureCall = callContext.iCS_S_VariableOrProcedureCall();
            if (variableOrProcedureCall != null)
            {
                return variableOrProcedureCall.ambiguousIdentifier();
            }

            return null;
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

            ResolveInternal(context, _currentScope);
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
                    type = _declarations.Where(d => d.IdentifierName == collection.GetText()).SingleOrDefault(item => item.IsBuiltIn && item.DeclarationType == DeclarationType.Class);
                    reference = CreateReference(baseType, type);
                }
            }
            else
            {
                type = ResolveType(asType.complexType());
                reference = CreateReference(asType.complexType(), type);
            }

            if (type != null && reference != null)
            {
                type.AddReference(reference);
                _alreadyResolved.Add(reference.Context);
            }
        }

        public void Resolve(VBAParser.ForNextStmtContext context)
        {
            var identifiers = context.ambiguousIdentifier();
            var identifier = ResolveInternal(identifiers[0], _currentScope, ContextAccessorType.AssignValue, null, false, true);
            if (identifier == null)
            {
                return;
            }

            // each iteration counts as an assignment
            var assignmentReference = CreateReference(identifiers[0], identifier, true);
            identifier.AddReference(assignmentReference);

            // each iteration also counts as a plain usage
            var usageReference = CreateReference(identifiers[0], identifier);
            identifier.AddReference(usageReference);

            if (identifiers.Count > 1)
            {
                var endForBlockReference = CreateReference(identifiers[1], identifier);
                identifier.AddReference(endForBlockReference);
            }
        }

        public void Resolve(VBAParser.ForEachStmtContext context)
        {
            var identifiers = context.ambiguousIdentifier();
            var identifier = ResolveInternal(identifiers[0], _currentScope, ContextAccessorType.AssignValue, null, false, true);
            if (identifier == null)
            {
                return;
            }

            // each iteration counts as an assignment
            var assignmentReference = CreateReference(identifiers[0], identifier, true);
            identifier.AddReference(assignmentReference);

            // each iteration also counts as a plain usage
            var usageReference = CreateReference(identifiers[0], identifier);
            identifier.AddReference(usageReference);

            if (identifiers.Count > 1)
            {
                identifier.AddReference(CreateReference(identifiers[1], identifier));
            }
        }

        public void Resolve(VBAParser.ImplementsStmtContext context)
        {
            ResolveInternal(context.ambiguousIdentifier(), _currentScope);
        }

        public void Resolve(VBAParser.RaiseEventStmtContext context)
        {
            ResolveInternal(context.ambiguousIdentifier(), _currentScope);
        }

        public void Resolve(VBAParser.ResumeStmtContext context)
        {
            ResolveInternal(context.ambiguousIdentifier(), _currentScope);
        }

        public void Resolve(VBAParser.FileNumberContext context)
        {
            ResolveInternal(context.ambiguousIdentifier(), _currentScope);
        }

        public void Resolve(VBAParser.ArgDefaultValueContext context)
        {
            ResolveInternal(context.ambiguousIdentifier(), _currentScope);
        }

        public void Resolve(VBAParser.FieldLengthContext context)
        {
            ResolveInternal(context.ambiguousIdentifier(), _currentScope);
        }

        public void Resolve(VBAParser.VsAssignContext context)
        {
            // named parameter reference must be scoped to called procedure
            var callee = FindParentCall(context);
            ResolveInternal(context.implicitCallStmt_InStmt(), callee, ContextAccessorType.GetValueOrReference);
        }

        private Declaration FindParentCall(VBAParser.VsAssignContext context)
        {
            var calleeContext = context.Parent.Parent.Parent;
            return ResolveInternal(calleeContext as VBAParser.ICS_B_ProcedureCallContext)
                   ?? ResolveInternal(calleeContext as VBAParser.ICS_S_VariableOrProcedureCallContext, _currentScope)
                   ?? ResolveInternal(calleeContext as VBAParser.ICS_S_ProcedureOrArrayCallContext, _currentScope)
                   ?? ResolveInternal(calleeContext as VBAParser.ICS_S_MembersCallContext, _currentScope);
        }

        private Declaration FindFunctionOrPropertyGetter(string identifierName, Declaration localScope = null)
        {
            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var matches = _declarations.Where(d => d.IdentifierName == identifierName);
            var parent = matches.SingleOrDefault(item =>
                item.Scope == localScope.Scope);

            return parent;
        }

        private Declaration FindLocalScopeDeclaration(string identifierName, Declaration localScope = null, bool parentContextIsVariableOrProcedureCall = false, bool isAssignmentTarget= false)
        {
            if (localScope == null)
            {
                localScope = _currentScope;
            }

            if (_moduleTypes.Contains(localScope.DeclarationType))
            {
                // "local scope" is not intended to be module level.
                return null;
            }

            var matches = _declarations.Where(d => d.IdentifierName == identifierName);

            try
            {
                var results = matches.Where(item =>
                    (item.ParentScope == localScope.Scope || (isAssignmentTarget && item.Scope == localScope.Scope))
                    && localScope.Context.GetSelection().Contains(item.Selection)
                    && !_moduleTypes.Contains(item.DeclarationType))
                    .ToList();

                if (results.Count > 1 && isAssignmentTarget
                    && _returningMemberTypes.Contains(localScope.DeclarationType)
                    && localScope.IdentifierName == identifierName
                    && parentContextIsVariableOrProcedureCall)
                {
                    // if we have multiple matches and we're in a returning member,
                    // in an in-statement variable or procedure call context that's
                    // the target of an assignment, then we have to assume we're looking
                    // at the assignment of the member's return value, i.e.:
                    /*
                     *    Property Get Foo() As Integer
                     *        Foo = 42 '<~ this Foo here
                     *    End Sub
                     */
                    return FindFunctionOrPropertyGetter(identifierName, localScope);
                }

                // if we're not returning a function/getter value, then there can be only one:
                return results.SingleOrDefault(item => !item.Equals(localScope));
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        private Declaration FindModuleScopeDeclaration(string identifierName, Declaration localScope = null)
        {
            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var matches = _declarations.Where(d => d.IdentifierName == identifierName);
            try
            {
                return matches.SingleOrDefault(item =>
                    item.ParentScope == localScope.ParentScope
                    && !item.DeclarationType.HasFlag(DeclarationType.Member)
                    && !_moduleTypes.Contains(item.DeclarationType)
                    && (item.DeclarationType != DeclarationType.Event || IsLocalEvent(item, localScope)));
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        private bool IsLocalEvent(Declaration item, Declaration localScope)
        {
            return item.DeclarationType == DeclarationType.Event
                   && localScope.Project == _currentScope.Project
                   && localScope.ComponentName == _currentScope.ComponentName;
        }

        private Declaration FindModuleScopeProcedure(string identifierName, Declaration localScope, ContextAccessorType accessorType, bool isAssignmentTarget = false)
        {
            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var matches = _declarations.Where(d => d.IdentifierName == identifierName);
            try
            {
                return matches.SingleOrDefault(item =>
                    item.Project == localScope.Project 
                    && item.ComponentName == localScope.ComponentName 
                    && (IsProcedure(item, localScope) || IsPropertyAccessor(item, accessorType, localScope, isAssignmentTarget)));
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        private Declaration FindProjectScopeDeclaration(string identifierName, Declaration localScope = null, bool hasStringQualifier = false)
        {
            // the "$" in e.g. "UCase$" isn't picked up as part of the identifierName, so we need to add it manually:
            var matches = _declarations.Where(item => !item.IsBuiltIn && item.IdentifierName == identifierName
                || item.IdentifierName == identifierName + (hasStringQualifier ? "$" : string.Empty)).ToList();

            if (matches.Count == 1)
            {
                return matches.Single();
            }

            if (localScope == null && _withBlockQualifiers.Any())
            {
                localScope = _withBlockQualifiers.Peek();
            }

            try
            {
                return SingleOrDefault(matches, IsUserDeclarationInProjectScope)
                    ?? SingleOrDefault(matches, item => IsBuiltInDeclarationInScope(item, localScope));
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        /// <summary>
        /// Returns a <see cref="Declaration"/> if exactly one match is found, <c>null</c> otherwise.
        /// </summary>
        private static TSource SingleOrDefault<TSource>(IEnumerable<TSource> source, Func<TSource, bool> predicate) where TSource : Declaration
        {
            try
            {
                var matches = source.Where(predicate).ToArray();
                return !matches.Any()
                    ? null
                    : matches.Single();
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        private static bool IsPublicOrGlobal(Declaration item)
        {
            return item.Accessibility == Accessibility.Global
                || item.Accessibility == Accessibility.Public;
        }

        private bool IsUserDeclarationInProjectScope(Declaration item)
        {
            var isNonMemberUserDeclaration = !item.IsBuiltIn 
                && !item.DeclarationType.HasFlag(DeclarationType.Member)
                // events can't be called outside the class they're declared in, exclude them as well:
                && item.DeclarationType != DeclarationType.Event;

            // declaration is in-scope if it's public/global, or if it's a module/class:
            return isNonMemberUserDeclaration && (IsPublicOrGlobal(item) || _moduleTypes.Contains(item.DeclarationType));
        }

        private static bool IsBuiltInDeclarationInScope(Declaration item, Declaration localScope)
        {
            var isBuiltInNonEvent = item.IsBuiltIn && item.DeclarationType != DeclarationType.Event;
            
            // if localScope is null, we can only resolve to a global:
            // note: built-in declarations are designed that way
            var isBuiltInGlobal = localScope == null && item.Accessibility == Accessibility.Global;

            // if localScope is not null, we can resolve to any public or global in that scope:
            var isInLocalScope = localScope != null && IsPublicOrGlobal(item)
                && localScope.IdentifierName == item.ParentDeclaration.IdentifierName;

            return isBuiltInNonEvent && (isBuiltInGlobal || isInLocalScope);
        }

        private static bool IsProcedure(Declaration item, Declaration localScope)
        {
            var isProcedure = item.DeclarationType == DeclarationType.Procedure
                              || item.DeclarationType == DeclarationType.Function;
            var isSameModule = item.Project == localScope.Project
                               && item.ComponentName == localScope.ComponentName;
            return isProcedure && isSameModule;
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