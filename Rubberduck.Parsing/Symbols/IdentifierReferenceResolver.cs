using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReferenceResolver
    {
        private readonly DeclarationFinder _declarationFinder;

        private enum ContextAccessorType
        {
            GetValueOrReference,
            AssignValue,
            AssignReference,
            AssignValueOrReference = AssignValue | AssignReference
        }

        private readonly QualifiedModuleName _qualifiedModuleName;

        private readonly IReadOnlyList<DeclarationType> _moduleTypes;
        private readonly IReadOnlyList<DeclarationType> _memberTypes;
        private readonly IReadOnlyList<DeclarationType> _returningMemberTypes;

        private readonly Stack<Declaration> _withBlockQualifiers;
        private readonly Stack<IBoundExpression> _withBlockExpressions;
        private readonly HashSet<RuleContext> _alreadyResolved;

        private readonly Declaration _moduleDeclaration;

        private Declaration _currentScope;
        private Declaration _currentParent;

        private readonly BindingService _bindingService;
        private readonly BoundExpressionVisitor _boundExpressionVisitor;

        public IdentifierReferenceResolver(QualifiedModuleName qualifiedModuleName, DeclarationFinder finder)
        {
            _declarationFinder = finder;

            _qualifiedModuleName = qualifiedModuleName;

            _withBlockQualifiers = new Stack<Declaration>();
            _withBlockExpressions = new Stack<IBoundExpression>();
            _alreadyResolved = new HashSet<RuleContext>();

            _moduleTypes = new[]
            {
                DeclarationType.ProceduralModule,
                DeclarationType.ClassModule,
            };

            _memberTypes = new[]
            {
                DeclarationType.Procedure,
                DeclarationType.Function,
                DeclarationType.PropertyGet,
                DeclarationType.PropertyLet,
                DeclarationType.PropertySet,
            };

            _returningMemberTypes = new[]
            {
                DeclarationType.Function,
                DeclarationType.PropertyGet,
            };

            _moduleDeclaration = finder.MatchName(_qualifiedModuleName.ComponentName)
                .SingleOrDefault(item =>
                    (item.DeclarationType == DeclarationType.ClassModule || item.DeclarationType == DeclarationType.ProceduralModule)
                && item.QualifiedName.QualifiedModuleName.Equals(_qualifiedModuleName));

            SetCurrentScope();

            var typeBindingContext = new TypeBindingContext(_declarationFinder);
            var procedurePointerBindingContext = new ProcedurePointerBindingContext(_declarationFinder);
            _bindingService = new BindingService(
                _declarationFinder,
                new DefaultBindingContext(_declarationFinder, typeBindingContext, procedurePointerBindingContext),
                typeBindingContext,
                procedurePointerBindingContext);
            _boundExpressionVisitor = new BoundExpressionVisitor();
        }

        public void SetCurrentScope()
        {
            _currentScope = _moduleDeclaration;
            _currentParent = _moduleDeclaration;
            _alreadyResolved.Clear();
        }

        public void SetCurrentScope(string memberName, DeclarationType type)
        {
            Debug.WriteLine("Setting current scope: {0} ({1}) in thread {2}", memberName, type, Thread.CurrentThread.ManagedThreadId);

            _currentParent = _declarationFinder.MatchName(memberName).SingleOrDefault(item =>
                item.QualifiedName.QualifiedModuleName == _qualifiedModuleName && item.DeclarationType == type);

            _currentScope = _declarationFinder.MatchName(memberName).SingleOrDefault(item =>
                item.QualifiedName.QualifiedModuleName == _qualifiedModuleName && item.DeclarationType == type) ?? _moduleDeclaration;

            Debug.WriteLine("Current scope is now {0} in thread {1}", _currentScope == null ? "null" : _currentScope.IdentifierName, Thread.CurrentThread.ManagedThreadId);
        }

        public void EnterWithBlock(VBAParser.WithStmtContext context)
        {
            Declaration qualifier = null;
            var expr = context.withStmtExpression();
            var typeExpression = expr.GetText();
            var boundExpression = _bindingService.ResolveDefault(_moduleDeclaration, _currentParent, typeExpression, GetInnerMostWithExpression());
            if (boundExpression != null)
            {
                _boundExpressionVisitor.AddIdentifierReferences(boundExpression, declaration => CreateReference(expr, declaration));
                qualifier = boundExpression.ReferencedDeclaration;
            }
            Resolve(context.block());
            // note: pushes null if unresolved
            _withBlockQualifiers.Push(qualifier);
            _withBlockExpressions.Push(boundExpression);
        }

        private IBoundExpression GetInnerMostWithExpression()
        {
            if (_withBlockExpressions.Any())
            {
                return _withBlockExpressions.Peek();
            }
            return null;
        }

        public void ExitWithBlock()
        {
            _withBlockQualifiers.Pop();
            _withBlockExpressions.Pop();
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

        private IEnumerable<IAnnotation> FindAnnotations(int line)
        {
            var annotationAbove = _declarationFinder.ModuleAnnotations(_qualifiedModuleName).SingleOrDefault(annotation => annotation.QualifiedSelection.Selection.EndLine == line - 1);
            if (annotationAbove != null)
            {
                return new List<IAnnotation>()
                {
                    annotationAbove
                };
            }
            return new List<IAnnotation>();
        }

        private void ResolveType(VBAParser.ICS_S_MembersCallContext context)
        {
            //var first = context.iCS_S_VariableOrProcedureCall().identifier();
            //var identifiers = new[] { first }.Concat(context.iCS_S_MemberCall()
            //            .Select(member => member.iCS_S_VariableOrProcedureCallUnrestricted().unrestrictedIdentifier()))
            //            .ToList();
            //ResolveType(identifiers);
        }

        private Declaration ResolveType(VBAParser.ComplexTypeContext context)
        {
            if (context == null)
            {
                return null;
            }

            var identifiers = context.identifier()
                .Select(identifier => identifier)
                .ToList();

            // if there's only 1 identifier, resolve to the tightest-scope match:
            if (identifiers.Count == 1)
            {
                var type = ResolveInScopeType(identifiers.Single().GetText(), _currentScope);
                if (type != null && !_alreadyResolved.Contains(context))
                {
                    type.AddReference(CreateReference(context, type));
                    _alreadyResolved.Add(context);
                }
                return type;
            }

            // if there's 2 or more identifiers, resolve to the deepest path:
            return ResolveType(identifiers);
        }

        private Declaration ResolveType(IList<VBAParser.IdentifierContext> identifiers)
        {
            var first = identifiers[0].GetText();
            var projectMatch = _declarationFinder.FindProject(first, _currentScope);

            if (projectMatch != null)
            {
                var projectReference = CreateReference(identifiers[0], projectMatch);

                // matches current project. 2nd identifier could be:
                // - standard module (only if there's a 3rd identifier)
                // - class module
                // - UDT
                // - Enum
                if (identifiers.Count == 3)
                {
                    var moduleMatch = _declarationFinder.FindStdModule(identifiers[1].GetText(), _currentScope);
                    if (moduleMatch != null)
                    {
                        var moduleReference = CreateReference(identifiers[1], moduleMatch);

                        // 3rd identifier can only be a UDT
                        var udtMatch = _declarationFinder.FindUserDefinedType(identifiers[2].GetText(), moduleMatch);
                        if (udtMatch != null)
                        {
                            var udtReference = CreateReference(identifiers[2], udtMatch);

                            if (!_alreadyResolved.Contains(projectReference.Context))
                            {
                                projectMatch.AddReference(projectReference);
                                _alreadyResolved.Add(projectReference.Context);
                            }

                            if (!_alreadyResolved.Contains(moduleReference.Context))
                            {
                                moduleMatch.AddReference(moduleReference);
                                _alreadyResolved.Add(moduleReference.Context);
                            }

                            if (!_alreadyResolved.Contains(udtReference.Context))
                            {
                                udtMatch.AddReference(udtReference);
                                _alreadyResolved.Add(udtReference.Context);
                            }

                            return udtMatch;
                        }
                        var enumMatch = _declarationFinder.FindEnum(identifiers[2].GetText(), moduleMatch);
                        if (enumMatch != null)
                        {
                            var enumReference = CreateReference(identifiers[2], enumMatch);

                            if (!_alreadyResolved.Contains(projectReference.Context))
                            {
                                projectMatch.AddReference(projectReference);
                                _alreadyResolved.Add(projectReference.Context);
                            }

                            if (!_alreadyResolved.Contains(moduleReference.Context))
                            {
                                moduleMatch.AddReference(moduleReference);
                                _alreadyResolved.Add(moduleReference.Context);
                            }

                            if (!_alreadyResolved.Contains(enumReference.Context))
                            {
                                enumMatch.AddReference(enumReference);
                                _alreadyResolved.Add(enumReference.Context);
                            }

                            return enumMatch;
                        }
                    }
                }
                else
                {
                    if (projectReference != null && !_alreadyResolved.Contains(projectReference.Context))
                    {
                        projectMatch.AddReference(projectReference);
                        _alreadyResolved.Add(projectReference.Context);
                    }

                    var match = _declarationFinder.FindClass(projectMatch, identifiers[1].GetText())
                                ?? _declarationFinder.FindUserDefinedType(identifiers[1].GetText())
                                ?? _declarationFinder.FindEnum(identifiers[1].GetText());
                    if (match != null)
                    {
                        var reference = CreateReference(identifiers[1], match);
                        if (reference != null && !_alreadyResolved.Contains(reference.Context))
                        {
                            match.AddReference(reference);
                            _alreadyResolved.Add(reference.Context);
                        }
                        return match;
                    }
                }
            }

            // first identifier didn't match current project.
            // if there are 3 identifiers, type isn't in current project.
            if (identifiers.Count != 3)
            {

                var moduleMatch = _declarationFinder.FindStdModule(identifiers[0].GetText(), projectMatch);
                if (moduleMatch != null)
                {
                    var moduleReference = CreateReference(identifiers[0], moduleMatch);

                    // 2nd identifier can only be a UDT or enum
                    var match = _declarationFinder.FindUserDefinedType(identifiers[1].GetText(), moduleMatch)
                            ?? _declarationFinder.FindEnum(identifiers[1].GetText(), moduleMatch);
                    if (match != null)
                    {
                        var reference = CreateReference(identifiers[1], match);

                        if (!_alreadyResolved.Contains(moduleReference.Context))
                        {
                            moduleMatch.AddReference(moduleReference);
                            _alreadyResolved.Add(moduleReference.Context);
                        }

                        if (!_alreadyResolved.Contains(reference.Context))
                        {
                            match.AddReference(reference);
                            _alreadyResolved.Add(reference.Context);
                        }

                        return match;
                    }
                }
            }

            return null;
        }

        private Declaration ResolveInScopeType(string identifier, Declaration scope)
        {
            var matches = _declarationFinder.MatchTypeName(identifier).ToList();
            if (matches.Count == 1)
            {
                return matches.Single();
            }

            if (matches.Count(match => match.ProjectId == scope.ProjectId) == 1)
            {
                return matches.Single(match => match.ProjectId == scope.ProjectId);
            }

            // more than one matching identifiers found.
            // if it matches a UDT or enum in the current scope, resolve to that type.
            var sameScopeUdt = matches.Where(declaration =>
                declaration.ProjectId == scope.ProjectId
                && (declaration.DeclarationType == DeclarationType.UserDefinedType
                || declaration.DeclarationType == DeclarationType.Enumeration)
                && declaration.ParentDeclaration.Equals(scope))
                .ToList();

            if (sameScopeUdt.Count == 1)
            {
                return sameScopeUdt.Single();
            }
            return null;
        }

        private Declaration ResolveType(Declaration parent)
        {
            if (parent != null && (parent.DeclarationType == DeclarationType.UserDefinedType
                                || parent.DeclarationType == DeclarationType.Enumeration
                                || parent.DeclarationType == DeclarationType.Project
                                || parent.DeclarationType == DeclarationType.ProceduralModule
                                || (parent.DeclarationType == DeclarationType.ClassModule && (parent.IsBuiltIn || parent.HasPredeclaredId))))
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

            identifier = identifier.StartsWith("VT_") ? parent.IdentifierName : identifier;

            var matches = _declarationFinder.MatchTypeName(identifier).ToList();
            if (matches.Count == 1)
            {
                return matches.Single();
            }

            var result = matches.Where(item =>
                (item.DeclarationType == DeclarationType.UserDefinedType
                || item.DeclarationType == DeclarationType.Enumeration)
                && item.ProjectId == _currentScope.ProjectId
                && item.ComponentName == _currentScope.ComponentName)
            .ToList();

            if (!result.Any())
            {
                result = matches.Where(item =>
                    _moduleTypes.Contains(item.DeclarationType)
                    && item.ProjectId == _currentScope.ProjectId)
                .ToList();
            }

            if (!result.Any())
            {
                result = matches.Where(item =>
                    _moduleTypes.Contains(item.DeclarationType))
                .ToList();
            }

            return result.Count == 1 ? result.SingleOrDefault() :
                matches.Count == 1 ? matches.First() : null;
        }

        private static readonly Type[] IdentifierContexts =
        {
            typeof (VBAParser.IdentifierContext),
            typeof (VBAParser.UnrestrictedIdentifierContext),
        };

        private Declaration ResolveInternal(ParserRuleContext callSiteContext, Declaration localScope, ContextAccessorType accessorType = ContextAccessorType.GetValueOrReference, VBAParser.DictionaryCallStmtContext fieldCall = null, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (callSiteContext == null)
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
            if (localScope.DeclarationType == DeclarationType.UserDefinedType)
            {
                callee = _declarationFinder.MatchName(identifierName).SingleOrDefault(item => item.Context != null && item.Context.Parent == localScope.Context);
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
                      ?? FindProjectScopeDeclaration(identifierName, Equals(localScope, _currentScope) ? null : localScope, accessorType, hasStringQualifier);
            }

            if (callee == null)
            {
                return null;
            }

            var reference = CreateReference(callSiteContext, callee, isAssignmentTarget, hasExplicitLetStatement);
            if (reference != null && !_alreadyResolved.Contains(reference.Context))
            {
                callee.AddReference(reference);
                _alreadyResolved.Add(reference.Context);
                _alreadyResolved.Add(callSiteContext);
            }

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
                ?? FindProjectScopeDeclaration(identifierName, Equals(localScope, _currentScope) ? null : localScope, accessorType, hasStringQualifier);
        }

        private Declaration ResolveInternal(VBAParser.ICS_S_VariableOrProcedureCallContext context, Declaration localScope, ContextAccessorType accessorType = ContextAccessorType.GetValueOrReference, bool hasExplicitLetStatement = false, bool isAssignmentTarget = false)
        {
            if (context == null)
            {
                return null;
            }
            if (ParserRuleContextHelper.HasParent<VBAParser.ImplementsStmtContext>(context))
            {
                return null;
            }
            if (ParserRuleContextHelper.HasParent<VBAParser.VsAddressOfContext>(context))
            {
                return null;
            }

            var identifierContext = context.identifier();
            var fieldCall = context.dictionaryCallStmt();

            var result = ResolveInternal(identifierContext, localScope, accessorType, fieldCall, hasExplicitLetStatement, isAssignmentTarget);
            if (result != null && localScope != null /*&& !localScope.DeclarationType.HasFlag(DeclarationType.Member)*/)
            {
                var reference = CreateReference(context.identifier(), result, isAssignmentTarget);
                if (reference != null)
                {
                    result.AddReference(reference);
                    //localScope.AddMemberCall(reference);
                }
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

            var fieldName = fieldCall.unrestrictedIdentifier().GetText();
            var result = _declarationFinder.MatchName(fieldName).SingleOrDefault(declaration => declaration.ParentScope == parentType.Scope);
            if (result == null)
            {
                return null;
            }

            var identifierContext = fieldCall.unrestrictedIdentifier();
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

            var identifierContext = context.identifier();
            var fieldCall = context.dictionaryCallStmt();
            // todo: understand WTF [baseType] is doing in that grammar rule...

            if (localScope == null)
            {
                localScope = _currentScope;
            }

            var result = ResolveInternal(identifierContext, localScope, accessorType, fieldCall, hasExplicitLetStatement, isAssignmentTarget);
            if (result != null && !localScope.DeclarationType.HasFlag(DeclarationType.Member))
            {
                localScope.AddMemberCall(CreateReference(context.identifier(), result));
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
                //// if we're on the left side of an assignment, only the last memberCall is the assignment target.
                //var isLast = memberCall.Equals(lastCall);
                //var accessor = isLast
                //    ? accessorType
                //    : ContextAccessorType.GetValueOrReference;
                //var isTarget = isLast && isAssignmentTarget;

                //var parentType = ResolveType(parent);

                //var member = ResolveInternal(memberCall.iCS_S_ProcedureOrArrayCallUnrestricted(), parentType, accessor, hasExplicitLetStatement, isTarget)
                //             ?? ResolveInternal(memberCall.iCS_S_VariableOrProcedureCallUnrestricted(), parentType, accessor, hasExplicitLetStatement, isTarget);

                //if (member == null && parent != null)
                //{
                //    var parentComTypeName = GetParentComTypeName(parent);

                //    // if the member can't be found on the parentType, maybe we're looking at a document or form module?
                //    parentType = _declarationFinder.FindClass(_moduleDeclaration.ParentDeclaration, parentComTypeName);
                //    member = ResolveInternal(memberCall.iCS_S_ProcedureOrArrayCallUnrestricted(), parentType, accessor, hasExplicitLetStatement, isTarget)
                //                 ?? ResolveInternal(memberCall.iCS_S_VariableOrProcedureCallUnrestricted(), parentType, accessor, hasExplicitLetStatement, isTarget);
                //}

                //if (member == null)
                //{
                //    // if member still can't be found, it's hopeless
                //    return null;
                //}

                //var memberReference = CreateReference(GetMemberCallIdentifierContext(memberCall), parent);
                //member.AddMemberCall(memberReference);
                //parent = ResolveType(member);
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

            var identifierContext = context.identifier();
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
                              ?? ResolveInternal(context.unrestrictedIdentifier(), parentType);
                parentType = ResolveType(parentScope);
            }

            var identifierContext = context.unrestrictedIdentifier();
            Declaration member = null;
            if (parentType != null)
            {
                member = _declarationFinder
                    .MatchName(identifierContext.GetText())
                    .SingleOrDefault(item =>
                        item.QualifiedName.QualifiedModuleName == parentType.QualifiedName.QualifiedModuleName
                        && item.DeclarationType != DeclarationType.Event);
            }
            else
            {
                if (parentScope != null)
                {
                    var parentComTypeName = GetParentComTypeName(parentScope);

                    // if the member can't be found on the parentType, maybe we're looking at a document or form module?
                    parentType = _declarationFinder.FindClass(_moduleDeclaration.ParentDeclaration, parentComTypeName);
                    member = ResolveInternal(identifierContext, parentType);
                }
            }

            if (member != null)
            {
                var reference = CreateReference(identifierContext, member);
                if (reference != null)
                {
                    parentScope.AddMemberCall(CreateReference(context.unrestrictedIdentifier(), member));
                    member.AddReference(reference);
                    _alreadyResolved.Add(reference.Context);
                }
            }
            else
            {
                return;
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
                var identifierContext = ((dynamic)parent.Context).identifier() as VBAParser.IdentifierContext;

                var parentReference = CreateReference(identifierContext, parent);
                if (parentReference != null)
                {
                    parent.AddReference(parentReference);
                    _alreadyResolved.Add(parentReference.Context);
                }
            }

            if (parent == null)
            {


                return;
            }

            var expression = context.GetText();
            var boundExpression = _bindingService.ResolveDefault(_moduleDeclaration, _currentParent, expression, GetInnerMostWithExpression());


            var chainedCalls = context.iCS_S_MemberCall();
            foreach (var memberCall in chainedCalls)
            {
                //var notationToken = memberCall.children[0];
                //if (notationToken.GetText() == "!")
                //{
                //    // the memberCall is a shorthand reference to the type's default member.
                //    // since the reference isn't explicit, we don't need to care for it.
                //    // (and we couldn't handle it if we wanted to, since we aren't parsing member attributes)
                //    return;
                //}

                //var member = ResolveInternal(memberCall.iCS_S_ProcedureOrArrayCall(), parent)
                //          ?? ResolveInternal(memberCall.iCS_S_VariableOrProcedureCall(), parent);

                //if (member == null && parent != null)
                //{
                //    var parentComTypeName = GetParentComTypeName(parent);
                //    // if the member can't be found on the parentType, maybe we're looking at a document or form module?
                //    var parentType = _declarationFinder.FindClass(null, parentComTypeName);
                //    member = ResolveInternal(memberCall.iCS_S_ProcedureOrArrayCall(), parentType)
                //                    ?? ResolveInternal(memberCall.iCS_S_VariableOrProcedureCall(), parentType);
                //}

                //if (member == null)
                //{
                //    return;
                //}

                //member.AddReference(CreateReference(GetMemberCallIdentifierContext(memberCall), member));
                //parent = ResolveType(member);
            }

            var fieldCall = context.dictionaryCallStmt();
            if (fieldCall == null)
            {
                return;
            }

            ResolveInternal(fieldCall, parent);
            _alreadyResolved.Add(context);
        }

        public void Resolve(VBAParser.OnErrorStmtContext context)
        {
            if (context.valueStmt() == null)
            {
                return;
            }
            ResolveLabel(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.ErrorStmtContext context)
        {
            ResolveDefault(context.valueStmt(), context.valueStmt().GetText());
        }

        private void ResolveLabel(ParserRuleContext context, string label)
        {
            var labelDeclaration = _bindingService.ResolveGoTo(_currentParent, label);
            if (labelDeclaration != null)
            {
                labelDeclaration.AddReference(CreateReference(context, labelDeclaration));
            }
        }

        private void ResolveDefault(ParserRuleContext context, string expression, bool isAssignmentTarget = false, bool hasExplicitLetStatement = false)
        {
            var boundExpression = _bindingService.ResolveDefault(_moduleDeclaration, _currentParent, expression, GetInnerMostWithExpression());
            if (boundExpression != null)
            {
                _boundExpressionVisitor.AddIdentifierReferences(boundExpression, declaration => CreateReference(context, declaration, isAssignmentTarget, hasExplicitLetStatement));
            }
        }

        public void Resolve(VBAParser.GoToStmtContext context)
        {
            ResolveLabel(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.OnGoToStmtContext context)
        {
            ResolveDefault(context.valueStmt()[0], context.valueStmt()[0].GetText());
            for (int labelIndex = 1; labelIndex < context.valueStmt().Count; labelIndex++)
            {
                ResolveLabel(context.valueStmt()[labelIndex], context.valueStmt()[labelIndex].GetText());
            }
        }

        public void Resolve(VBAParser.GoSubStmtContext context)
        {
            ResolveLabel(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.OnGoSubStmtContext context)
        {
            ResolveDefault(context.valueStmt()[0], context.valueStmt()[0].GetText());
            for (int labelIndex = 1; labelIndex < context.valueStmt().Count; labelIndex++)
            {
                ResolveLabel(context.valueStmt()[labelIndex], context.valueStmt()[labelIndex].GetText());
            }
        }

        public void Resolve(VBAParser.RedimStmtContext context)
        {
            foreach (var redimStmt in context.redimSubStmt())
            {
                foreach (var dimSpec in redimStmt.subscripts().subscript())
                {
                    foreach (var expr in dimSpec.valueStmt())
                    {
                        ResolveDefault(expr, expr.GetText());
                    }
                }
            }
        }

        public void Resolve(VBAParser.BlockContext context)
        {
            if (context == null)
            {
                return;
            }
            foreach (var stmt in context.blockStmt())
            {
                Resolve(stmt);
            }
        }

        public void Resolve(VBAParser.BlockStmtContext context)
        {
            //if (context == null)
            //{
            //    return;
            //}
            //dynamic ctx = context;
            //Resolve(ctx);
        }

        public void Resolve(VBAParser.WhileWendStmtContext context)
        {
            ResolveDefault(context.valueStmt(), context.valueStmt().GetText());
            Resolve(context.block());
        }

        public void Resolve(VBAParser.DoLoopStmtContext context)
        {
            if (context.valueStmt() == null)
            {
                return;
            }
            ResolveDefault(context.valueStmt(), context.valueStmt().GetText());
            Resolve(context.block());
        }

        public void Resolve(VBAParser.BlockIfThenElseContext context)
        {
            ResolveDefault(context.ifBlockStmt().ifConditionStmt(), context.ifBlockStmt().ifConditionStmt().GetText());
            Resolve(context.ifBlockStmt().block());
            if (context.ifElseIfBlockStmt() != null)
            {
                foreach (var elseIfBlock in context.ifElseIfBlockStmt())
                {
                    ResolveDefault(elseIfBlock.ifConditionStmt(), elseIfBlock.ifConditionStmt().GetText());
                    Resolve(elseIfBlock.block());
                }
            }
            if (context.ifElseBlockStmt() != null)
            {
                Resolve(context.ifElseBlockStmt().block());
            }
        }

        public void Resolve(VBAParser.InlineIfThenElseContext context)
        {
            ResolveDefault(context.ifConditionStmt(), context.ifConditionStmt().GetText());
            foreach (var blockStmt in context.blockStmt())
            {
                Resolve(blockStmt);
            }
        }

        public void Resolve(VBAParser.SelectCaseStmtContext context)
        {
            ResolveDefault(context.valueStmt(), context.valueStmt().GetText());
            if (context.sC_Case() != null)
            {
                foreach (var caseClauseBlock in context.sC_Case())
                {
                    var caseClause = caseClauseBlock.sC_Cond();
                    if (caseClause is VBAParser.CaseCondSelectionContext)
                    {
                        foreach (var selectClause in ((VBAParser.CaseCondSelectionContext)caseClause).sC_Selection())
                        {
                            if (selectClause is VBAParser.CaseCondIsContext)
                            {
                                var ctx = (VBAParser.CaseCondIsContext)selectClause;
                                ResolveDefault(ctx.valueStmt(), ctx.valueStmt().GetText());
                            }
                            else if (selectClause is VBAParser.CaseCondToContext)
                            {
                                var ctx = (VBAParser.CaseCondToContext)selectClause;
                                ResolveDefault(ctx.valueStmt()[0], ctx.valueStmt()[0].GetText());
                                ResolveDefault(ctx.valueStmt()[0], ctx.valueStmt()[0].GetText());
                            }
                            else
                            {
                                var ctx = (VBAParser.CaseCondValueContext)selectClause;
                                ResolveDefault(ctx.valueStmt(), ctx.valueStmt().GetText());
                            }
                        }
                    }
                    Resolve(caseClauseBlock.block());
                }
            }
        }

        private string GetParentComTypeName(Declaration declaration)
        {
            if (declaration.QualifiedName.QualifiedModuleName.Component == null)
            {
                return string.Empty;
            }
            Property property;
            try
            {
                property = declaration.QualifiedName.QualifiedModuleName.Component.Properties.OfType<Property>().Where(p => p.Name == "Parent").FirstOrDefault();
            }
            catch
            {
                // TODO: Doesn't work in MS Access.
                return string.Empty;
            }
            if (property != null)
            {
                return ComHelper.GetTypeName(property.Object);
            }
            return string.Empty;
        }

        private ParserRuleContext GetMemberCallIdentifierContext(VBAParser.ICS_S_MemberCallContext callContext)
        {
            if (callContext == null)
            {
                return null;
            }

            var procedureOrArrayCall = callContext.iCS_S_ProcedureOrArrayCallUnrestricted();
            if (procedureOrArrayCall != null)
            {
                return procedureOrArrayCall.unrestrictedIdentifier();
            }

            var variableOrProcedureCall = callContext.iCS_S_VariableOrProcedureCallUnrestricted();
            if (variableOrProcedureCall != null)
            {
                return variableOrProcedureCall.unrestrictedIdentifier();
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
            var letStatement = context.LET();
            ResolveDefault(context.valueStmt()[0], context.valueStmt()[0].GetText(), true, letStatement != null);
            ResolveDefault(context.valueStmt()[1], context.valueStmt()[1].GetText());
        }

        public void Resolve(VBAParser.SetStmtContext context)
        {
            ResolveDefault(context.valueStmt()[0], context.valueStmt()[0].GetText(), true, false);
            ResolveDefault(context.valueStmt()[1], context.valueStmt()[1].GetText());
        }

        public void Resolve(VBAParser.ExplicitCallStmtContext context)
        {
            ResolveDefault(context.explicitCallStmtExpression(), context.explicitCallStmtExpression().GetText());
        }

        public void Resolve(VBAParser.ConstStmtContext context)
        {
            foreach (var constStmt in context.constSubStmt())
            {
                ResolveDefault(constStmt.valueStmt(), constStmt.valueStmt().GetText());
            }
        }

        public void Resolve(VBAParser.EraseStmtContext context)
        {
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
        }

        public void Resolve(VBAParser.OpenStmtContext context)
        {
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
            ResolveDefault(context.fileNumber().valueStmt(), context.fileNumber().valueStmt().GetText());
        }

        public void Resolve(VBAParser.CloseStmtContext context)
        {
            foreach (var expr in context.fileNumber())
            {
                ResolveDefault(expr.valueStmt(), expr.valueStmt().GetText());
            }
        }

        public void Resolve(VBAParser.SeekStmtContext context)
        {
            ResolveDefault(context.fileNumber().valueStmt(), context.fileNumber().valueStmt().GetText());
            ResolveDefault(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.LockStmtContext context)
        {
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
        }

        public void Resolve(VBAParser.UnlockStmtContext context)
        {
            ResolveDefault(context.fileNumber().valueStmt(), context.fileNumber().valueStmt().GetText());
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
        }

        public void Resolve(VBAParser.LineInputStmtContext context)
        {
            ResolveDefault(context.fileNumber().valueStmt(), context.fileNumber().valueStmt().GetText());
            ResolveDefault(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.WidthStmtContext context)
        {
            ResolveDefault(context.fileNumber().valueStmt(), context.fileNumber().valueStmt().GetText());
            ResolveDefault(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.PrintStmtContext context)
        {
            ResolveDefault(context.fileNumber().valueStmt(), context.fileNumber().valueStmt().GetText());
            foreach (var expr in context.outputList().outputList_Expression())
            {
                if (expr.valueStmt() != null)
                {
                    ResolveDefault(expr.valueStmt(), expr.valueStmt().GetText());
                }
                ResolveArgsCall(expr.argsCall());
            }
        }

        public void Resolve(VBAParser.WriteStmtContext context)
        {
            ResolveDefault(context.fileNumber().valueStmt(), context.fileNumber().valueStmt().GetText());
            foreach (var expr in context.outputList().outputList_Expression())
            {
                if (expr.valueStmt() != null)
                {
                    ResolveDefault(expr.valueStmt(), expr.valueStmt().GetText());
                }
                ResolveArgsCall(expr.argsCall());
            }
        }

        public void Resolve(VBAParser.InputStmtContext context)
        {
            ResolveDefault(context.fileNumber().valueStmt(), context.fileNumber().valueStmt().GetText());
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
        }

        public void Resolve(VBAParser.PutStmtContext context)
        {
            ResolveDefault(context.fileNumber().valueStmt(), context.fileNumber().valueStmt().GetText());
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
        }

        public void Resolve(VBAParser.GetStmtContext context)
        {
            ResolveDefault(context.fileNumber().valueStmt(), context.fileNumber().valueStmt().GetText());
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
        }

        public void Resolve(VBAParser.LsetStmtContext context)
        {
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
        }

        public void Resolve(VBAParser.RsetStmtContext context)
        {
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
        }

        public void Resolve(VBAParser.AsTypeClauseContext context)
        {
            var asType = context.type();
            if (asType == null)
            {
                return;
            }
            var baseType = asType.baseType();
            if (baseType != null)
            {
                // Fixed-Length strings can have a constant-name as length that is a simple-name-expression that also has to be resolved.
                var length = context.fieldLength();
                if (context.fieldLength() != null && context.fieldLength().identifier() != null)
                {
                    var constantName = context.fieldLength().identifier();
                    var constantNameExpression = _bindingService.ResolveDefault(_moduleDeclaration, _currentParent, constantName.GetText(), GetInnerMostWithExpression());
                    if (constantNameExpression != null)
                    {
                        _boundExpressionVisitor.AddIdentifierReferences(constantNameExpression, declaration => CreateReference(constantName, declaration));
                    }
                }
                return;
            }
            string typeExpression = asType.complexType().GetText();
            var boundExpression = _bindingService.ResolveType(_moduleDeclaration, _currentParent, typeExpression);
            if (boundExpression != null)
            {
                _boundExpressionVisitor.AddIdentifierReferences(boundExpression, declaration => CreateReference(asType.complexType(), declaration));
            }
        }

        public void Resolve(VBAParser.ForNextStmtContext context)
        {
            var firstExpression = _bindingService.ResolveDefault(_moduleDeclaration, _currentParent, context.valueStmt()[0].GetText(), GetInnerMostWithExpression());
            if (firstExpression != null)
            {
                // each iteration counts as an assignment
                _boundExpressionVisitor.AddIdentifierReferences(firstExpression, declaration => CreateReference(context.valueStmt()[0], declaration, true));
                // each iteration also counts as a plain usage
                _boundExpressionVisitor.AddIdentifierReferences(firstExpression, declaration => CreateReference(context.valueStmt()[0], declaration));
            }
            for (int exprIndex = 1; exprIndex < context.valueStmt().Count; exprIndex++)
            {
                var expr = _bindingService.ResolveDefault(_moduleDeclaration, _currentParent, context.valueStmt()[exprIndex].GetText(), GetInnerMostWithExpression());
                if (expr != null)
                {
                    _boundExpressionVisitor.AddIdentifierReferences(expr, declaration => CreateReference(context.valueStmt()[exprIndex], declaration));
                }
            }
            Resolve(context.block());
        }

        public void Resolve(VBAParser.ForEachStmtContext context)
        {
            var firstExpression = _bindingService.ResolveDefault(_moduleDeclaration, _currentParent, context.valueStmt()[0].GetText(), GetInnerMostWithExpression());
            if (firstExpression != null)
            {
                // each iteration counts as an assignment
                _boundExpressionVisitor.AddIdentifierReferences(firstExpression, declaration => CreateReference(context.valueStmt()[0], declaration, true));
                // each iteration also counts as a plain usage
                _boundExpressionVisitor.AddIdentifierReferences(firstExpression, declaration => CreateReference(context.valueStmt()[0], declaration));
            }

            for (int exprIndex = 1; exprIndex < context.valueStmt().Count; exprIndex++)
            {
                var expr = _bindingService.ResolveDefault(_moduleDeclaration, _currentParent, context.valueStmt()[exprIndex].GetText(), GetInnerMostWithExpression());
                if (expr != null)
                {
                    _boundExpressionVisitor.AddIdentifierReferences(expr, declaration => CreateReference(context.valueStmt()[exprIndex], declaration));
                }
            }
            Resolve(context.block());
        }

        public void Resolve(VBAParser.ImplementsStmtContext context)
        {
            var boundExpression = _bindingService.ResolveType(_moduleDeclaration, _currentParent, context.valueStmt().GetText());
            if (boundExpression != null)
            {
                _boundExpressionVisitor.AddIdentifierReferences(boundExpression, declaration => CreateReference(context.valueStmt(), declaration));
            }
        }

        public void Resolve(VBAParser.VsAddressOfContext context)
        {
            var boundExpression = _bindingService.ResolveProcedurePointer(_moduleDeclaration, _currentParent, context.valueStmt().GetText());
            if (boundExpression != null)
            {
                _boundExpressionVisitor.AddIdentifierReferences(boundExpression, declaration => CreateReference(context.valueStmt(), declaration));
            }
        }

        public void Resolve(VBAParser.RaiseEventStmtContext context)
        {
            var eventDeclaration = _bindingService.ResolveEvent(_moduleDeclaration, context.identifier().GetText());
            if (eventDeclaration != null)
            {
                eventDeclaration.AddReference(CreateReference(context.identifier(), eventDeclaration));
            }
            ResolveArgsCall(context.argsCall());
        }

        public void Resolve(VBAParser.MidStmtContext context)
        {
            ResolveArgsCall(context.argsCall());
        }

        private void ResolveArgsCall(VBAParser.ArgsCallContext argsCall)
        {
            if (argsCall == null)
            {
                return;
            }
            foreach (var argCall in argsCall.argCall())
            {
                ResolveDefault(argCall.valueStmt(), argCall.valueStmt().GetText());
            }
        }

        public void Resolve(VBAParser.ResumeStmtContext context)
        {
            if (context.valueStmt() == null)
            {
                return;
            }
            ResolveLabel(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.LineLabelContext context)
        {
            // Nothing to bind.
        }

        public void Resolve(VBAParser.AttributeStmtContext context)
        {
            // Nothing to bind.
        }

        public void Resolve(VBAParser.DeftypeStmtContext context)
        {
            // Nothing to bind.
        }

        public void Resolve(VBAParser.ExitStmtContext context)
        {
            // Nothing to bind.
        }

        public void Resolve(VBAParser.VariableStmtContext context)
        {
            // Nothing to bind.
        }

        public void Resolve(VBAParser.ImplicitCallStmt_InBlockContext context)
        {
            // TODO: This is a call statement but arg list has to be specified separately.
        }

        public void Resolve(VBAParser.FieldLengthContext context)
        {
            ResolveInternal(context.identifier(), _currentScope);
        }

        public void Resolve(VBAParser.VsAssignContext context)
        {
            // TODO: Understand this.
            // named parameter reference must be scoped to called procedure
            var callee = FindParentCall(context);
            ResolveInternal(context.implicitCallStmt_InStmt(), callee, ContextAccessorType.AssignValueOrReference);
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

            var matches = _declarationFinder.MatchName(identifierName);
            var parent = matches.SingleOrDefault(item =>
                (item.DeclarationType.HasFlag(DeclarationType.Function) || item.DeclarationType.HasFlag(DeclarationType.PropertyGet))
                && item.Equals(localScope));

            return parent;
        }

        private Declaration FindLocalScopeDeclaration(string identifierName, Declaration localScope = null, bool parentContextIsVariableOrProcedureCall = false, bool isAssignmentTarget = false)
        {
            if (localScope == null)
            {
                localScope = _currentScope;
            }

            if (_moduleTypes.Contains(localScope.DeclarationType) || localScope.DeclarationType.HasFlag(DeclarationType.Project))
            {
                // "local scope" is not intended to be module level.
                return null;
            }

            var matches = _declarationFinder.MatchName(identifierName);

            var results = matches.Where(item =>
                ((localScope.Equals(item.ParentDeclaration)
                || (item.DeclarationType.HasFlag(DeclarationType.Parameter) && localScope.Equals(item.ParentScopeDeclaration)))
                || (isAssignmentTarget && item.Scope == localScope.Scope))
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

            if (localScope.DeclarationType.HasFlag(DeclarationType.Project))
            {
                return null;
            }

            if (identifierName == "Me" && _moduleDeclaration.DeclarationType.HasFlag(DeclarationType.ClassModule))
            {
                return _moduleDeclaration;
            }

            var scope = localScope; // avoid implicitly capturing 'this'
            var matches = _declarationFinder.MatchName(identifierName).Where(item => !item.Equals(scope)).ToList();

            var result = matches.Where(item =>
                (localScope.ParentScopeDeclaration == null || localScope.ParentScopeDeclaration.Equals(item.ParentScopeDeclaration))
                && !item.DeclarationType.HasFlag(DeclarationType.Member)
                && !_moduleTypes.Contains(item.DeclarationType)
                && item.DeclarationType != DeclarationType.UserDefinedType && item.DeclarationType != DeclarationType.Enumeration
                && (item.DeclarationType != DeclarationType.Event || IsLocalEvent(item, localScope)))
            .ToList();

            if (matches.Any() && !result.Any())
            {
                result = matches.Where(item =>
                    (localScope != null && localScope.Equals(item.ParentScopeDeclaration))
                    && !item.DeclarationType.HasFlag(DeclarationType.Member)
                    && !_moduleTypes.Contains(item.DeclarationType)
                    && item.DeclarationType != DeclarationType.UserDefinedType && item.DeclarationType != DeclarationType.Enumeration
                    && (item.DeclarationType != DeclarationType.Event || IsLocalEvent(item, localScope)))
                .ToList();
            }

            return result.Count == 1 ? result.SingleOrDefault() : null; // return null for multiple matches
        }

        private bool IsLocalEvent(Declaration item, Declaration localScope)
        {
            return item.DeclarationType == DeclarationType.Event
                   && localScope.ProjectId == _currentScope.ProjectId
                   && localScope.ComponentName == _currentScope.ComponentName;
        }

        private Declaration FindModuleScopeProcedure(string identifierName, Declaration localScope, ContextAccessorType accessorType, bool isAssignmentTarget = false)
        {
            if (localScope == null)
            {
                localScope = _currentScope;
            }

            if (localScope.DeclarationType == DeclarationType.Project)
            {
                return null;
            }

            var matches = _declarationFinder.MatchName(identifierName);
            var result = matches.Where(item =>
                _memberTypes.Contains(item.DeclarationType)
                && localScope.ProjectId == item.ProjectId
                && (localScope.ComponentName.Replace("_", string.Empty) == item.ComponentName.Replace("_", string.Empty))
                && (IsProcedure(item, localScope) || IsPropertyAccessor(item, accessorType, localScope, isAssignmentTarget)))
            .ToList();

            return result.Count == 1 ? result.SingleOrDefault() : null;
        }

        private bool IsStdModuleMember(Declaration declaration)
        {
            return declaration.ParentDeclaration != null
                   && declaration.ParentDeclaration.DeclarationType == DeclarationType.ProceduralModule;
        }

        private bool IsPublicEnum(Declaration declaration)
        {
            return (IsPublicOrGlobal(declaration) || declaration.Accessibility == Accessibility.Implicit)
                   && (declaration.DeclarationType == DeclarationType.Enumeration
                       || declaration.DeclarationType == DeclarationType.EnumerationMember);
        }

        private bool IsStaticClass(Declaration declaration)
        {
            var isDocumentOrForm = !declaration.IsBuiltIn &&
                (declaration.QualifiedName.QualifiedModuleName.Component.Type == vbext_ComponentType.vbext_ct_Document
                ||
                declaration.QualifiedName.QualifiedModuleName.Component.Type == vbext_ComponentType.vbext_ct_MSForm);

            return isDocumentOrForm || (declaration.ParentDeclaration != null
                   && declaration.ParentDeclaration.DeclarationType == DeclarationType.ClassModule
                   && declaration.ParentDeclaration.HasPredeclaredId);

        }

        private readonly IReadOnlyList<string> SpecialCasedTokens = new[]{
            Tokens.Error,
            Tokens.Hex,
            Tokens.Oct,
            Tokens.Str,
            Tokens.CurDir,
            Tokens.Command,
            Tokens.Environ,
            Tokens.Chr,
            Tokens.ChrW,
            Tokens.Format,
            Tokens.LCase,
            Tokens.Left,
            Tokens.LeftB,
            Tokens.LTrim,
            Tokens.Mid,
            Tokens.MidB,
            Tokens.Trim,
            Tokens.Right,
            Tokens.RightB,
            Tokens.RTrim,
            Tokens.UCase
        };

        private Declaration FindProjectScopeDeclaration(string identifierName, Declaration localScope = null, ContextAccessorType accessorType = ContextAccessorType.GetValueOrReference, bool hasStringQualifier = false)
        {
            var matches = _declarationFinder.MatchName(identifierName).Where(item =>
                item.DeclarationType == DeclarationType.Project
                || item.DeclarationType == DeclarationType.ProceduralModule
                || IsPublicEnum(item)
                || IsStaticClass(item)
                || IsStdModuleMember(item)
                || (item.ParentScopeDeclaration != null && item.ParentScopeDeclaration.Equals(localScope))).ToList();

            if (matches.Count == 1 && !SpecialCasedTokens.Contains(matches.Single().IdentifierName))
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
            if (result.Count == 1 && !SpecialCasedTokens.Contains(result.Single().IdentifierName))
            {
                return result.SingleOrDefault();
            }
            else
            {
                var nonModules = matches.Where(item => !_moduleTypes.Contains(item.DeclarationType)).ToList();
                var temp = nonModules.Where(item => item.DeclarationType ==
                                                    (accessorType == ContextAccessorType.GetValueOrReference
                                                        ? DeclarationType.PropertyGet
                                                        : item.DeclarationType))
                    .ToList();
                if (temp.Count > 1)
                {
                    if (localScope == null)
                    {
                        var names = new[] { "Global", "_Global" };
                        var appGlobals = temp.Where(item => names.Contains(item.ParentDeclaration.IdentifierName)).ToList();
                        if (appGlobals.Count == 1)
                        {
                            return appGlobals.Single();
                        }
                    }
                    else
                    {
                        var names = new[] { localScope.IdentifierName, "I" + localScope.IdentifierName };
                        var members = temp.Where(item => names.Contains(item.ParentScopeDeclaration.IdentifierName)
                                                         && item.DeclarationType == (accessorType == ContextAccessorType.GetValueOrReference
                                                             ? DeclarationType.PropertyGet
                                                             : item.DeclarationType)).ToList();
                        if (members.Count == 1)
                        {
                            return members.Single();
                        }
                    }

                    Debug.WriteLine("Ambiguous match in '{0}': '{1}'", localScope == null ? "(unknown)" : localScope.IdentifierName, identifierName);
                }
            }

            // VBA.Strings.Left function is actually called _B_var_Left;
            // VBA.Strings.Left$ is _B_str_Left.
            // same for all $-terminated functions.
            var surrogateName = hasStringQualifier
                ? "_B_str_" + identifierName
                : "_B_var_" + identifierName;

            matches = _declarationFinder.MatchName(surrogateName).ToList();
            if (matches.Count == 1)
            {
                return matches.Single();
            }

            Debug.WriteLine("Unknown identifier in '{0}': '{1}'", localScope == null ? "(unknown)" : localScope.IdentifierName, identifierName);
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
            var isInLocalScope = (localScope != null && item.Accessibility == Accessibility.Global
                && localScope.IdentifierName == item.ParentDeclaration.IdentifierName)
                || (localScope != null && localScope.QualifiedName.QualifiedModuleName.Component != null
                    && localScope.QualifiedName.QualifiedModuleName.Component.Type == Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document
                 && item.Accessibility == Accessibility.Public && item.ParentDeclaration.DeclarationType == localScope.DeclarationType);

            return isBuiltInNonEvent && (isBuiltInGlobal || isInLocalScope);
        }

        private static bool IsProcedure(Declaration item, Declaration localScope)
        {
            var isProcedure = item.DeclarationType == DeclarationType.Procedure
                              || item.DeclarationType == DeclarationType.Function;
            var isSameModule = item.ProjectId == localScope.ProjectId
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