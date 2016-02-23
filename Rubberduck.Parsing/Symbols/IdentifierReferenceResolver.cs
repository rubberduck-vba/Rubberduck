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

        private readonly IReadOnlyList<Declaration> _declarations;
        private readonly IReadOnlyList<CommentNode> _comments;

        private readonly QualifiedModuleName _qualifiedModuleName;

        private readonly IReadOnlyList<DeclarationType> _moduleTypes;
        private readonly IReadOnlyList<DeclarationType> _scopingTypes;
        private readonly IReadOnlyList<DeclarationType> _parentTypes;
        private readonly IReadOnlyList<DeclarationType> _returningMemberTypes;

        private readonly IReadOnlyList<Accessibility> _projectScopePublicModifiers; 

        private readonly Stack<Declaration> _withBlockQualifiers;
        private readonly HashSet<RuleContext> _alreadyResolved;

        private readonly Declaration _moduleDeclaration;
        private readonly IReadOnlyList<Declaration> _scopingDeclarations;
        private readonly IReadOnlyList<Declaration> _parentDeclarations;

        private Declaration _currentScope;
        private Declaration _currentParent;

        public IdentifierReferenceResolver(QualifiedModuleName qualifiedModuleName, IReadOnlyList<Declaration> declarations, IReadOnlyList<CommentNode> comments)
        {
            _qualifiedModuleName = qualifiedModuleName;
            _declarations = declarations;
            _comments = comments;

            _withBlockQualifiers = new Stack<Declaration>();
            _alreadyResolved = new HashSet<RuleContext>();

            _moduleTypes = new[]
            {
                DeclarationType.Module, 
                DeclarationType.Class,
            };

            _scopingTypes =new[]
            {
                DeclarationType.Function, 
                DeclarationType.Procedure, 
                DeclarationType.PropertyGet, 
                DeclarationType.PropertyLet, 
                DeclarationType.PropertySet,
            };

            _parentTypes = new[]
            {
                DeclarationType.Function, 
                DeclarationType.Procedure, 
                DeclarationType.PropertyGet, 
                DeclarationType.PropertyLet, 
                DeclarationType.PropertySet,
                DeclarationType.Enumeration, 
                DeclarationType.UserDefinedType, 
            };

            _returningMemberTypes = new[]
            {
                DeclarationType.Function,
                DeclarationType.PropertyGet, 
            };

            _projectScopePublicModifiers = new[]
            {
                Accessibility.Public, 
                Accessibility.Global, 
                Accessibility.Friend, 
                Accessibility.Implicit, 
            };

            _moduleDeclaration = _declarations.SingleOrDefault(item =>
                _moduleTypes.Contains(item.DeclarationType)
                && item.Project == _qualifiedModuleName.Project
                && item.ComponentName == _qualifiedModuleName.ComponentName);

            _scopingDeclarations = _declarations.Where(item =>
                _scopingTypes.Contains(item.DeclarationType)
                && item.Project == _qualifiedModuleName.Project
                && item.ComponentName == _qualifiedModuleName.ComponentName).ToList();

            _parentDeclarations = _declarations.Where(item =>
                _parentTypes.Contains(item.DeclarationType)
                && item.Project == _qualifiedModuleName.Project
                && item.ComponentName == _qualifiedModuleName.ComponentName).ToList();

            SetCurrentScope();
        }

        public void SetCurrentScope()
        {
            _currentScope = _moduleDeclaration;
            _currentParent = _moduleDeclaration;
            _alreadyResolved.Clear();
        }

        public void SetCurrentScope(string memberName, DeclarationType type)
        {
            _currentParent = _parentDeclarations.SingleOrDefault(item =>
                item.DeclarationType == type && item.IdentifierName == memberName);

            _currentScope = _scopingDeclarations.SingleOrDefault(item =>
                item.DeclarationType == type && item.IdentifierName == memberName) ?? _moduleDeclaration;
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
            if (callSiteContext == null || _currentScope == null || _alreadyResolved.Contains(callSiteContext))
            {
                return null;
            }
            var name = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var annotations = FindAnnotations(selection.StartLine);
            return new IdentifierReference(_qualifiedModuleName, _currentScope, _currentParent, name, selection, callSiteContext, callee, isAssignmentTarget, hasExplicitLetStatement, annotations);
        }

        private string FindAnnotations(int line)
        {
            if (_comments == null)
            {
                return null;
            }

            var commentAbove = _comments.SingleOrDefault(comment => comment.QualifiedSelection.QualifiedName == _qualifiedModuleName && comment.QualifiedSelection.Selection.EndLine == line - 1);
            if (commentAbove != null && commentAbove.CommentText.StartsWith("@"))
            {
                return commentAbove.CommentText;
            }
            return null;
        }

        private void ResolveType(VBAParser.ICS_S_MembersCallContext context)
        {
            var first = context.iCS_S_VariableOrProcedureCall().ambiguousIdentifier();
            var identifiers = new[] {first}.Concat(context.iCS_S_MemberCall()
                        .Select(member => member.iCS_S_VariableOrProcedureCall().ambiguousIdentifier()))
                        .ToList();
            ResolveType(identifiers);
        }

        private void ResolveType(VBAParser.ComplexTypeContext context)
        {
            if (context == null)
            {
                return;
            }

            var identifiers = context.ambiguousIdentifier()
                .Select(identifier => identifier)
                .ToList();

            // if there's only 1 identifier, resolve to the tightest-scope match:
            if (identifiers.Count == 1)
            {
                var type = ResolveInScopeType(identifiers.Single().GetText(), _currentScope);
                if (type != null)
                {
                    type.AddReference(CreateReference(context, type));
                    _alreadyResolved.Add(context);
                }
                return;
            }

            // if there's 2 or more identifiers, resolve to the deepest path:
            ResolveType(identifiers);
        }

        private void ResolveType(IList<VBAParser.AmbiguousIdentifierContext> identifiers)
        {
            var first = identifiers[0].GetText();
            var projectMatch = _currentScope.ProjectName == first
                ? _declarations.SingleOrDefault(declaration =>
                    declaration.DeclarationType == DeclarationType.Project
                    && declaration.Project == _currentScope.Project // todo: account for project references!
                    && declaration.IdentifierName == first)
                : null;

            if (projectMatch != null)
            {
                var projectReference = CreateReference(identifiers[0], projectMatch);

                // matches current project. 2nd identifier could be:
                // - standard module (only if there's a 3rd identifier)
                // - class module
                // - UDT
                if (identifiers.Count == 3)
                {
                    var moduleMatch = _declarations.SingleOrDefault(declaration =>
                        !declaration.IsBuiltIn && declaration.ParentDeclaration != null
                        && declaration.ParentDeclaration.Equals(projectMatch)
                        && declaration.DeclarationType == DeclarationType.Module
                        && declaration.IdentifierName == identifiers[1].GetText());

                    if (moduleMatch != null)
                    {
                        var moduleReference = CreateReference(identifiers[1], moduleMatch);

                        // 3rd identifier can only be a UDT
                        var udtMatch = _declarations.SingleOrDefault(declaration =>
                            !declaration.IsBuiltIn && declaration.ParentDeclaration != null
                            && declaration.ParentDeclaration.Equals(moduleMatch)
                            && declaration.DeclarationType == DeclarationType.UserDefinedType
                            && declaration.IdentifierName == identifiers[2].GetText());
                        if (udtMatch != null)
                        {
                            var udtReference = CreateReference(identifiers[2], udtMatch);

                            projectMatch.AddReference(projectReference);
                            _alreadyResolved.Add(projectReference.Context);

                            moduleMatch.AddReference(moduleReference);
                            _alreadyResolved.Add(moduleReference.Context);

                            udtMatch.AddReference(udtReference);
                            _alreadyResolved.Add(udtReference.Context);

                            return;
                        }
                    }
                }
                else
                {
                    projectMatch.AddReference(projectReference);
                    _alreadyResolved.Add(projectReference.Context);

                    var match = _declarations.SingleOrDefault(declaration =>
                        !declaration.IsBuiltIn && declaration.ParentDeclaration != null
                        && declaration.ParentDeclaration.Equals(projectMatch)
                        && declaration.IdentifierName == identifiers[1].GetText()
                        && (declaration.DeclarationType == DeclarationType.Class ||
                            declaration.DeclarationType == DeclarationType.UserDefinedType));
                    if (match != null)
                    {
                        var reference = CreateReference(identifiers[1], match);
                        if (reference != null)
                        {
                            match.AddReference(reference);
                            _alreadyResolved.Add(reference.Context);
                            return;
                        }
                    }
                }
            }

            // first identifier didn't match current project.
            // if there are 3 identifiers, type isn't in current project.
            if (identifiers.Count != 3)
            {
                var moduleMatch = _declarations.SingleOrDefault(declaration =>
                    !declaration.IsBuiltIn && declaration.ParentDeclaration != null
                    && declaration.ParentDeclaration.Equals(projectMatch)
                    && declaration.DeclarationType == DeclarationType.Module
                    && declaration.IdentifierName == identifiers[0].GetText());

                if (moduleMatch != null)
                {
                    var moduleReference = CreateReference(identifiers[0], moduleMatch);

                    // 2nd identifier can only be a UDT
                    var udtMatch = _declarations.SingleOrDefault(declaration =>
                        !declaration.IsBuiltIn && declaration.ParentDeclaration != null
                        && declaration.ParentDeclaration.Equals(moduleMatch)
                        && declaration.DeclarationType == DeclarationType.UserDefinedType
                        && declaration.IdentifierName == identifiers[1].GetText());
                    if (udtMatch != null)
                    {
                        var udtReference = CreateReference(identifiers[1], udtMatch);

                        moduleMatch.AddReference(moduleReference);
                        _alreadyResolved.Add(moduleReference.Context);

                        udtMatch.AddReference(udtReference);
                        _alreadyResolved.Add(udtReference.Context);
                    }
                }
            }
        }

        private IEnumerable<Declaration> FindMatchingTypes(string identifier)
        {
            return _declarations.Where(declaration =>
                declaration.IdentifierName == identifier
                && (declaration.DeclarationType == DeclarationType.Class
                || declaration.DeclarationType == DeclarationType.UserDefinedType))
                .ToList();
        }

        private Declaration ResolveInScopeType(string identifier, Declaration scope)
        {
            var matches = FindMatchingTypes(identifier).ToList();
            if (matches.Count == 1)
            {
                return matches.Single();
            }

            // more than one matching identifiers found.
            // if it matches a UDT in the current scope, resolve to that type.
            var sameScopeUdt = matches.Where(declaration =>
                declaration.Project == scope.Project
                && declaration.DeclarationType == DeclarationType.UserDefinedType
                && declaration.ParentDeclaration.Equals(scope))
                .ToList();

            if (sameScopeUdt.Count == 1)
            {
                return sameScopeUdt.Single();
            }
            
            // todo: try to resolve identifier using referenced projects

            return null;
        }




        private Declaration ResolveType(Declaration parent)
        {
            if (parent != null && parent.DeclarationType == DeclarationType.UserDefinedType)
            {
                return parent;
            }

            if (parent == null || parent.AsTypeName == null)
            {
                return null;
            }

            var identifier = parent.AsTypeName.Contains(".")
                ? parent.AsTypeName.Split('.').Last() // bug: this can't be right
                : parent.AsTypeName;

            var matches = _declarations.Where(d => d.IdentifierName == identifier).ToList();

            var result = matches.Where(item =>
                item.DeclarationType == DeclarationType.UserDefinedType
                && item.Project == _currentScope.Project
                && item.ComponentName == _currentScope.ComponentName)
            .ToList();

            if (!result.Any())
            {
                result = matches.Where(item =>
                    _moduleTypes.Contains(item.DeclarationType)
                    && item.Project == _currentScope.Project)
                .ToList();                
            }

            if (!result.Any())
            {
                result = matches.Where(item =>
                    _moduleTypes.Contains(item.DeclarationType))
                .ToList();
            }

            return result.Count == 1 ? result.SingleOrDefault() : null;
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
            if (identifierName.StartsWith("[") && identifierName.EndsWith("]"))
            {
                // square-bracketed identifier may contain a '!' symbol; identifier name is at the left of it.
                identifierName = identifierName.Substring(1, identifierName.Length - 2)/*.Split('!').First()*/;
                // problem, is that IdentifierReference should work off IDENTIFIER tokens, not AmbiguousIdentifierContext.
                // not sure what the better fix is. 
            }

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
                var reference = CreateReference(context.ambiguousIdentifier(), result, isAssignmentTarget);
                result.AddReference(reference);
                localScope.AddMemberCall(reference);
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
                if (parent == null)
                {
                    // if parent is an unknown type, continuing any further will only cause issues.
                    return null;
                }
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
            if (reference != null)
            {
                callee.AddReference(reference);
                _alreadyResolved.Add(reference.Context);
            }
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

            if (context.Parent.Parent.Parent is VBAParser.VsNewContext)
            {
                // if we're in a ValueStatement/New context, we're actually resolving for a type:
                ResolveType(context);
                return;
            }

            Declaration parent;
            if (_withBlockQualifiers.Any())
            {
                parent = ResolveType(_withBlockQualifiers.Peek());
                if (parent == null)
                {
                    return;
                }
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
                if (parentReference != null)
                {
                    parent.AddReference(parentReference);
                    _alreadyResolved.Add(parentReference.Context);
                }
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
                ResolveType(asType.complexType());
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

            // each iteration also counts as a plain usage - CreateReference will return null here, need to create it manually.
            var name = identifiers[0].GetText();
            var selection = identifiers[0].GetSelection();
            var annotations = FindAnnotations(selection.StartLine);
            var usageReference = new IdentifierReference(_qualifiedModuleName, _currentScope, _currentParent, name, selection, identifiers[0], identifier, false, false, annotations);
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
                item.ParentScopeDeclaration.Equals(localScope));

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

            var results = matches.Where(item =>
                (item.ParentScope == localScope.Scope || (isAssignmentTarget && item.Scope == localScope.Scope))
                && localScope.Context.GetSelection().Contains(item.Selection)
                && !_moduleTypes.Contains(item.DeclarationType))
                .ToList();

            if (results.Count >= 1 && isAssignmentTarget
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
            var result = results.Where(item => !item.Equals(localScope)).ToList();
            return result.Count == 1 ? result.SingleOrDefault() : null;
        }

        private Declaration FindModuleScopeDeclaration(string identifierName, Declaration localScope = null)
        {
            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var matches = _declarations.Where(d => d.IdentifierName == identifierName);
            var result = matches.Where(item =>
                item.ParentScope == localScope.ParentScope
                && !item.DeclarationType.HasFlag(DeclarationType.Member)
                && !_moduleTypes.Contains(item.DeclarationType)
                && (item.DeclarationType != DeclarationType.Event || IsLocalEvent(item, localScope)))
            .ToList();

            return result.Count == 1 ? result.SingleOrDefault() : null;
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
            var result = matches.Where(item =>
                item.Project == localScope.Project 
                && item.ComponentName == localScope.ComponentName 
                && (IsProcedure(item, localScope) || IsPropertyAccessor(item, accessorType, localScope, isAssignmentTarget)))
            .ToList();

            return result.Count == 1 ? result.SingleOrDefault() : null;
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

            var result = matches.Where(IsUserDeclarationInProjectScope).ToList();
            if (result.Count == 1)
            {
                return result.SingleOrDefault();
            }

            result = matches.Where(item => IsBuiltInDeclarationInScope(item, localScope)).ToList();
            if (result.Count == 1)
            {
                return result.SingleOrDefault();
            }

            return null;
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