using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReferenceListener : VBABaseListener
    {
        private readonly Declarations _declarations;
        private readonly QualifiedModuleName _qualifiedName;

        private string _currentScope;
        private DeclarationType _currentScopeType;

        public IdentifierReferenceListener(VBComponentParseResult result, Declarations declarations)
            : this(result.QualifiedName, declarations)
        { }

        public IdentifierReferenceListener(QualifiedModuleName qualifiedName, Declarations declarations)
        {
            _qualifiedName = qualifiedName;
            _declarations = declarations;
            SetCurrentScope();
        }

        private string ModuleScope { get { return _qualifiedName.ToString(); } }

        /// <summary>
        /// Sets current scope to module-level.
        /// </summary>
        private void SetCurrentScope()
        {
            _currentScope = ModuleScope;
            _currentScopeType = _qualifiedName.Component.Type == vbext_ComponentType.vbext_ct_StdModule
                ? DeclarationType.Module
                : DeclarationType.Class;
        }

        /// <summary>
        /// Sets current scope to specified module member.
        /// </summary>
        private void SetCurrentScope(string name, DeclarationType scopeType)
        {
            _currentScope = _qualifiedName + "." + name;
            _currentScopeType = scopeType;
        }

        public override void EnterLiteral(VBAParser.LiteralContext context)
        {
            var stringLiteral = context.STRINGLITERAL();
            if (stringLiteral != null)
            {
                HandleEmptyStringLiteral(stringLiteral);
                return;
            }

            var numberLiteral = context.INTEGERLITERAL();
            if (numberLiteral != null)
            {
                HandleNumberLiteral(numberLiteral);
                return;
            }
        }

        private void HandleEmptyStringLiteral(ITerminalNode stringLiteral)
        {
            if (stringLiteral.Symbol.Text.Length == 2) // string literal only contains opening & closing quotes
            {
                // todo: track that value + implement an inspection that recommends replacing it with vbNullString (#363)
            }
        }

        private void HandleNumberLiteral(ITerminalNode numberLiteral)
        {
            // todo: verify whether the string representation ends with a type hint; flag as such if needed.

            // todo: track that value + implement an inspection that checks for magic numbers (#359)
            // also, don't do anything here if tree walker is currently in a ConstSubStmtContext

            
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.Procedure);
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText(), DeclarationType.Function);
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

        public override void EnterLetStmt(VBAParser.LetStmtContext context)
        {
            var leftSide = context.implicitCallStmt_InStmt();
            var letStatement = context.LET();
            var target = FindAssignmentTarget(leftSide, DeclarationType.PropertyLet);
            if (target != null)
            {
                EnterIdentifier(target, target.GetSelection(), true, letStatement != null);
            }
        }

        public override void EnterSetStmt(VBAParser.SetStmtContext context)
        {
            var leftSide = context.implicitCallStmt_InStmt();
            var target = FindAssignmentTarget(leftSide, DeclarationType.PropertySet);
            if (target != null)
            {
                EnterIdentifier(target, target.GetSelection(), true);
            }
        }

        private VBAParser.AmbiguousIdentifierContext FindAssignmentTarget(VBAParser.ImplicitCallStmt_InStmtContext leftSide, DeclarationType accessorType)
        {
            VBAParser.AmbiguousIdentifierContext context;
            var call = Resolve(leftSide.iCS_S_ProcedureOrArrayCall(), out context, accessorType)
                       ?? Resolve(leftSide.iCS_S_VariableOrProcedureCall(), out context, accessorType)
                       ?? Resolve(leftSide.iCS_S_DictionaryCall(), out context, accessorType)
                       ?? Resolve(leftSide.iCS_S_MembersCall(), out context, accessorType);

            return context;
        }

        private VBAParser.AmbiguousIdentifierContext EnterDictionaryCall(VBAParser.DictionaryCallStmtContext dictionaryCall, VBAParser.AmbiguousIdentifierContext parentIdentifier = null, DeclarationType accessorType = DeclarationType.PropertyGet)
        {
            if (dictionaryCall == null)
            {
                return null;
            }

            if (parentIdentifier != null)
            {
                var isTarget = accessorType == DeclarationType.PropertyLet || accessorType == DeclarationType.PropertySet;
                if (!EnterIdentifier(parentIdentifier, parentIdentifier.GetSelection(), isTarget, accessorType:accessorType))
                    // we're referencing "member" in "member!field"
                {
                    return null;
                }
            }

            var identifier = dictionaryCall.ambiguousIdentifier();
            if (_declarations.Items.Any(item => item.IdentifierName == identifier.GetText()))
            {
                return identifier;
            }

            return null;
        }

        public override void EnterComplexType(VBAParser.ComplexTypeContext context)
        {
            var identifiers = context.ambiguousIdentifier();
            _skipIdentifiers = !identifiers.All(identifier => _declarations.Items.Any(declaration => declaration.IdentifierName == identifier.GetText()));
        }

        public override void ExitComplexType(VBAParser.ComplexTypeContext context)
        {
            _skipIdentifiers = false;
        }

        private bool _skipIdentifiers;
        public override void EnterAmbiguousIdentifier(VBAParser.AmbiguousIdentifierContext context)
        {
            if (_skipIdentifiers || IsDeclarativeContext(context))
            {
                return;
            }

            var selection = context.GetSelection();

            if (IsAssignmentContext(context))
            {
                EnterIdentifier(context, selection, true);
            }
            else
            {
                EnterIdentifier(context, selection);
            }
        }

        private bool IsAssignmentContext(ParserRuleContext context)
        {
            return context.Parent is VBAParser.ForNextStmtContext
                   || context.Parent is VBAParser.ForEachStmtContext
                   || context.Parent.Parent.Parent.Parent is VBAParser.LineInputStmtContext
                   || context.Parent.Parent.Parent.Parent is VBAParser.InputStmtContext;
        }

        public override void EnterCertainIdentifier(VBAParser.CertainIdentifierContext context)
        {
            // skip declarations
            if (IsDeclarativeContext(context))
            {
                return;
            }

            var selection = context.GetSelection();
            EnterIdentifier(context, selection);
        }

        private bool EnterIdentifier(ParserRuleContext context, Selection selection, bool isAssignmentTarget = false, bool hasExplicitLetStatement = false, DeclarationType accessorType = DeclarationType.PropertyGet)
        {
            var name = context.GetText();
            var matches = _declarations[name].Where(IsInScope);

            var declaration = GetClosestScopeDeclaration(matches, context, accessorType);
            if (declaration != null)
            {
                var reference = new IdentifierReference(_qualifiedName, name, selection, context, declaration, isAssignmentTarget, hasExplicitLetStatement);

                // thread-local copy
                var references = declaration.References.ToList();
                if (!references.Select(r => r.Context).Contains(reference.Context))
                {
                    declaration.AddReference(reference);
                    return true;
                }
                // note: non-matching names are not necessarily undeclared identifiers, e.g. "String" in "Dim foo As String".
            }

            return false;
        }

        public override void EnterVsNew(VBAParser.VsNewContext context)
        {
            _skipIdentifiers = true;
            var identifiers = context.valueStmt().GetRuleContexts<VBAParser.ImplicitCallStmt_InStmtContext>();

            var lastIdentifier = identifiers.Last();
            var name = lastIdentifier.GetText();

            var matches = _declarations[name].Where(d => d.DeclarationType == DeclarationType.Class).ToList();
            var result = matches.Count <= 1 
                ? matches.SingleOrDefault()
                : GetClosestScopeDeclaration(matches, context, DeclarationType.Class);

            if (result == null)
            {
                return;
            }

            var reference = new IdentifierReference(_qualifiedName, result.IdentifierName, lastIdentifier.GetSelection(), context, result);
            result.AddReference(reference);
        }

        public override void ExitVsNew(VBAParser.VsNewContext context)
        {
            _skipIdentifiers = false;
        }

        private readonly Stack<Declaration> _withQualifiers = new Stack<Declaration>();
        public override void EnterWithStmt(VBAParser.WithStmtContext context)
        {
            var implicitCall = context.implicitCallStmt_InStmt();

            var call = Resolve(implicitCall.iCS_S_ProcedureOrArrayCall())
                ?? Resolve(implicitCall.iCS_S_VariableOrProcedureCall())
                ?? Resolve(implicitCall.iCS_S_DictionaryCall())
                ?? Resolve(implicitCall.iCS_S_MembersCall());

            _withQualifiers.Push(GetReturnType(call));            
        }

        private Declaration GetReturnType(Declaration call)
        {
            return call == null 
                ? null 
                : _declarations.Items.SingleOrDefault(item =>
                item.DeclarationType == DeclarationType.Class
                && item.Accessibility != Accessibility.Private
                && item.IdentifierName == call.AsTypeName);
        }

        public override void ExitWithStmt(VBAParser.WithStmtContext context)
        {
            _withQualifiers.Pop();
        }

        private Declaration Resolve(VBAParser.ICS_S_ProcedureOrArrayCallContext context, out VBAParser.AmbiguousIdentifierContext identifierContext, DeclarationType accessorType)
        {
            if (context == null)
            {
                identifierContext = null;
                return null;
            }

            var identifier = context.ambiguousIdentifier();
            var name = identifier.GetText();

            var procedure = FindProcedureDeclaration(name, identifier);
            var result = procedure ?? FindVariableDeclaration(name, identifier, accessorType);

            identifierContext = result == null 
                ? null 
                : result.Context == null 
                    ? null 
                    : ((dynamic) result.Context).ambiguousIdentifier();
            return result;
        }

        private Declaration Resolve(VBAParser.ICS_S_ProcedureOrArrayCallContext context)
        {
            VBAParser.AmbiguousIdentifierContext discarded;
            return Resolve(context, out discarded, DeclarationType.PropertyGet);
        }

        private Declaration Resolve(VBAParser.ICS_S_VariableOrProcedureCallContext context, out VBAParser.AmbiguousIdentifierContext identifierContext, DeclarationType accessorType)
        {
            if (context == null)
            {
                identifierContext = null;
                return null;
            }

            var identifier = context.ambiguousIdentifier();
            var name = identifier.GetText();

            var procedure = FindProcedureDeclaration(name, identifier, accessorType);
            var result = procedure ?? FindVariableDeclaration(name, identifier, accessorType);

            identifierContext = result == null 
                ? null 
                : result.Context == null 
                    ? null 
                    : ((dynamic) result.Context).ambiguousIdentifier();
            return result;
        }

        private Declaration Resolve(VBAParser.ICS_S_VariableOrProcedureCallContext context)
        {
            VBAParser.AmbiguousIdentifierContext discarded;
            return Resolve(context, out discarded, DeclarationType.PropertyGet);
        }

        private Declaration Resolve(VBAParser.ICS_S_DictionaryCallContext context, out VBAParser.AmbiguousIdentifierContext identifierContext, DeclarationType accessorType, VBAParser.AmbiguousIdentifierContext parentIdentifier = null)
        {
            if (context == null)
            {
                identifierContext = null;
                return null;
            }

            var identifier = EnterDictionaryCall(context.dictionaryCallStmt(), parentIdentifier, accessorType);
            var name = identifier.GetText();

            var result = FindVariableDeclaration(name, identifier, accessorType);

            identifierContext = result == null 
                ? null 
                : result.Context == null 
                    ? null 
                    : ((dynamic) result.Context).ambiguousIdentifier();
            return result;
        }

        private Declaration Resolve(VBAParser.ICS_S_DictionaryCallContext context, VBAParser.AmbiguousIdentifierContext parentIdentifier = null)
        {
            VBAParser.AmbiguousIdentifierContext discarded;
            return Resolve(context, out discarded, DeclarationType.PropertyGet, parentIdentifier);
        }

        private Declaration Resolve(VBAParser.ICS_S_MembersCallContext context, out VBAParser.AmbiguousIdentifierContext identifierContext, DeclarationType accessorType)
        {
            if (context == null)
            {
                identifierContext = null;
                return null;
            }

            var members = context.iCS_S_MemberCall();
            for (var index = 0; index < members.Count; index++)
            {
                var member = members[index];
                if (index < members.Count - 1)
                {
                    var parent = Resolve(member.iCS_S_ProcedureOrArrayCall())
                                         ?? Resolve(member.iCS_S_VariableOrProcedureCall());

                    if (parent == null)
                    {
                        // return early if we can't resolve the whole member chain
                        identifierContext = null;
                        return null;
                    }
                }
                else
                {
                    var result = Resolve(member.iCS_S_ProcedureOrArrayCall())
                                 ?? Resolve(member.iCS_S_VariableOrProcedureCall());

                    identifierContext = result == null 
                        ? null 
                        : result.Context == null 
                            ? null 
                            : ((dynamic) result.Context).ambiguousIdentifier();
                    return result;
                }
            }

            identifierContext = null;
            return null;
        }

        private Declaration Resolve(VBAParser.ICS_S_MembersCallContext context)
        {
            VBAParser.AmbiguousIdentifierContext discarded;
            return Resolve(context, out discarded, DeclarationType.PropertyGet);
        }

        private Declaration Resolve(VBAParser.ICS_B_MemberProcedureCallContext context)
        {
            var parent = context.implicitCallStmt_InStmt();
            var parentCall = Resolve(parent.iCS_S_VariableOrProcedureCall())
                             ?? Resolve(parent.iCS_S_ProcedureOrArrayCall())
                             ?? Resolve(parent.iCS_S_DictionaryCall())
                             ?? Resolve(parent.iCS_S_MembersCall());

            if (parentCall == null)
            {
                return null;
            }

            var type = _declarations[parentCall.AsTypeName].SingleOrDefault(item =>
                item.DeclarationType == DeclarationType.Class
                || item.DeclarationType == DeclarationType.UserDefinedType);

            var members = _declarations.FindMembers(type);
            var name = context.ambiguousIdentifier().GetText();

            return members.SingleOrDefault(m => m.IdentifierName == name);
        }

        public override void EnterVsAssign(VBAParser.VsAssignContext context)
        {
            /* named parameter syntax */

            // one of these is null...
            var callStatementA = context.Parent.Parent.Parent as VBAParser.ICS_S_ProcedureOrArrayCallContext;
            var callStatementB = context.Parent.Parent.Parent as VBAParser.ICS_S_VariableOrProcedureCallContext;
            var callStatementC = context.Parent.Parent.Parent as VBAParser.ICS_B_MemberProcedureCallContext;
            var callStatementD = context.Parent.Parent.Parent as VBAParser.ICS_B_ProcedureCallContext;
            
            var procedureName = string.Empty;
            ParserRuleContext identifierContext = null;
            if (callStatementA != null)
            {
                procedureName = callStatementA.ambiguousIdentifier().GetText();
                identifierContext = callStatementA.ambiguousIdentifier();
            }
            else if(callStatementB != null)
            {
                procedureName = callStatementB.ambiguousIdentifier().GetText();
                identifierContext = callStatementB.ambiguousIdentifier();
            }
            else if (callStatementC != null)
            {
                procedureName = callStatementC.ambiguousIdentifier().GetText();
                identifierContext = callStatementC.ambiguousIdentifier();
            }
            else if (callStatementD != null)
            {
                procedureName = callStatementD.certainIdentifier().GetText();
                identifierContext = callStatementD.certainIdentifier();
            }

            var procedure = FindProcedureDeclaration(procedureName, identifierContext);
            if (procedure == null)
            {
                return;
            }

            var call = context.implicitCallStmt_InStmt();
            var arg = Resolve(call.iCS_S_VariableOrProcedureCall())
                      ?? Resolve(call.iCS_S_ProcedureOrArrayCall())
                      ?? Resolve(call.iCS_S_DictionaryCall())
                      ?? Resolve(call.iCS_S_MembersCall());

            if (arg != null)
            {
                var reference = new IdentifierReference(_qualifiedName, arg.IdentifierName, context.GetSelection(), context, arg);
                arg.AddReference(reference);
            }
        }

        private static readonly DeclarationType[] PropertyAccessors =
        {
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        private Declaration FindProcedureDeclaration(string procedureName, ParserRuleContext context, DeclarationType accessor = DeclarationType.PropertyGet)
        {
            var matches = _declarations[procedureName]
                .Where(declaration => ProcedureDeclarations.Contains(declaration.DeclarationType))
                .Where(IsInScope)
                .ToList();

            if (!matches.Any())
            {
                return null;
            }

            if (matches.Count == 1)
            {
                return matches.First();
            }

            if (matches.All(m => PropertyAccessors.Contains(m.DeclarationType)))
            {
                return matches.Find(m => m.DeclarationType == accessor);
            }

            var procedure = GetClosestScopeDeclaration(matches, context);
            return procedure;
        }

        private Declaration FindVariableDeclaration(string procedureName, ParserRuleContext context, DeclarationType accessorType)
        {
            var matches = _declarations[procedureName]
                .Where(declaration => declaration.DeclarationType == DeclarationType.Variable || declaration.DeclarationType == DeclarationType.Parameter)
                .Where(IsInScope);

            var variable = GetClosestScopeDeclaration(matches, context, accessorType);
            return variable;
        }

        private static readonly DeclarationType[] ProcedureDeclarations = 
            {
                DeclarationType.Procedure,
                DeclarationType.Function,
                DeclarationType.PropertyGet,
                DeclarationType.PropertyLet,
                DeclarationType.PropertySet
            };

        private bool IsInScope(Declaration declaration)
        {
            if (declaration.IsBuiltIn && declaration.Accessibility == Accessibility.Global)
            {
                return true; // global-scope built-in identifiers are always in scope
            }

            if (declaration.DeclarationType == DeclarationType.Project)
            {
                return true; // a project name is always in scope anywhere
            }

            if (declaration.DeclarationType == DeclarationType.Module ||
                declaration.DeclarationType == DeclarationType.Class)
            {
                // todo: access component instancing properties to do this right (class)
                // i.e. a private class in another project wouldn't be accessible
                return true;
            }

            if (ProcedureDeclarations.Contains(declaration.DeclarationType))
            {
                if (declaration.Accessibility == Accessibility.Public 
                 || declaration.Accessibility == Accessibility.Implicit)
                {
                    var result = _qualifiedName.Project.Equals(declaration.Project);
                    return result;
                }

                return declaration.QualifiedName.QualifiedModuleName == _qualifiedName;
            }

            return declaration.Scope == _currentScope
                   || declaration.Scope == ModuleScope
                   || IsGlobalField(declaration) 
                   || IsGlobalProcedure(declaration);
        }

        private static readonly Type[] PropertyContexts =
        {
            typeof (VBAParser.PropertyGetStmtContext),
            typeof (VBAParser.PropertyLetStmtContext),
            typeof (VBAParser.PropertySetStmtContext)
        };

        private Declaration GetClosestScopeDeclaration(IEnumerable<Declaration> declarations, ParserRuleContext context, DeclarationType accessorType = DeclarationType.PropertyGet)
        {
            if (context.Parent.Parent.Parent is VBAParser.AsTypeClauseContext)
            {
                accessorType = DeclarationType.Class;
            }

            var matches = declarations as IList<Declaration> ?? declarations.ToList();
            if (!matches.Any())
            {
                return null;
            }

            // handle indexed property getters
            var currentScopeMatches = matches.Where(declaration =>
                (declaration.Scope == _currentScope && !PropertyContexts.Contains(declaration.Context.Parent.Parent.GetType()))
                || ((declaration.Context != null && declaration.Context.Parent.Parent is VBAParser.PropertyGetStmtContext
                    && _currentScopeType == DeclarationType.PropertyGet)
                || (declaration.Context != null && declaration.Context.Parent.Parent is VBAParser.PropertySetStmtContext
                    && _currentScopeType == DeclarationType.PropertySet)
                || (declaration.Context != null && declaration.Context.Parent.Parent is VBAParser.PropertyLetStmtContext
                    && _currentScopeType == DeclarationType.PropertyLet)))
                .ToList();
            if (currentScopeMatches.Count == 1)
            {
                return currentScopeMatches[0];
            }

            // note: commented-out because it breaks the UDT member references, but property getters behave strangely still
            //var currentScope = matches.SingleOrDefault(declaration =>
            //    IsCurrentScopeMember(accessorType, declaration)
            //    && (declaration.DeclarationType == accessorType
            //        || accessorType == DeclarationType.PropertyGet));

            //if (matches.First().IdentifierName == "procedure")
            //{
            //    // for debugging - "procedure" is both a UDT member and a parameter to a procedure.
            //}

            if (matches.Count == 1)
            {
                return matches[0];
            }

            var moduleScope = matches.SingleOrDefault(declaration => declaration.Scope == ModuleScope);
            if (moduleScope != null)
            {
                return moduleScope;
            }

            var splitScope = _currentScope.Split('.');
            if (splitScope.Length > 2) // Project.Module.Procedure - i.e. if scope is deeper than module-level
            {
                var scope = splitScope[0] + '.' + splitScope[1];
                var scopeMatches = matches.Where(m => m.ParentScope == scope
                                                      && (!PropertyAccessors.Contains(m.DeclarationType)
                                                          || m.DeclarationType == accessorType)).ToList();
                if (scopeMatches.Count == 1)
                {
                    return scopeMatches.Single();
                }

                // handle standard library member shadowing:
                if (!matches.All(m => m.IsBuiltIn))
                {
                    var ambiguousMatches = matches.Where(m => !m.IsBuiltIn
                                                              && (!PropertyAccessors.Contains(m.DeclarationType)
                                                                  || m.DeclarationType == accessorType)).ToList();

                    if (ambiguousMatches.Count == 1)
                    {
                        return ambiguousMatches.Single();
                    }
                }
            }

            var memberProcedureCallContext = context.Parent as VBAParser.ICS_B_MemberProcedureCallContext;
            if (memberProcedureCallContext != null)
            {
                return Resolve(memberProcedureCallContext);
            }

            var implicitCall = context.Parent.Parent as VBAParser.ImplicitCallStmt_InStmtContext;
            if (implicitCall != null)
            {
                return Resolve(implicitCall.iCS_S_VariableOrProcedureCall())
                       ?? Resolve(implicitCall.iCS_S_ProcedureOrArrayCall())
                       ?? Resolve(implicitCall.iCS_S_DictionaryCall())
                       ?? Resolve(implicitCall.iCS_S_MembersCall());
            }

            return null;
        }

        private bool IsCurrentScopeMember(DeclarationType accessorType, Declaration declaration)
        {
            if (declaration.Scope != ModuleScope && accessorType != DeclarationType.Class)
            {
                return false;
            }

            switch (accessorType)
            {
                case DeclarationType.Class:
                    return declaration.DeclarationType == DeclarationType.Class;

                case DeclarationType.PropertySet:
                    return declaration.DeclarationType != DeclarationType.PropertyGet && declaration.DeclarationType != DeclarationType.PropertyLet;

                case DeclarationType.PropertyLet:
                    return declaration.DeclarationType != DeclarationType.PropertyGet && declaration.DeclarationType != DeclarationType.PropertySet;

                case DeclarationType.PropertyGet:
                    return declaration.DeclarationType != DeclarationType.PropertyLet && declaration.DeclarationType != DeclarationType.PropertySet;

                default:
                    return true;
            } 
        }

        private bool IsGlobalField(Declaration declaration)
        {
            // a field isn't a field if it's not a variable or a constant.
            if (declaration.DeclarationType != DeclarationType.Variable ||
                declaration.DeclarationType != DeclarationType.Constant)
            {
                return false;
            }

            // a field is only global if it's declared as Public or Global in a standard module.
            var moduleMatches = _declarations[declaration.ComponentName].ToList();
            var modules = moduleMatches.Where(match => match.DeclarationType == DeclarationType.Module);

            // Friend members are only visible within the same project.
            var isSameProject = declaration.Project == _qualifiedName.Project;

            // todo: verify that this isn't overkill. Friend modifier has restricted legal usage.
            return modules.Any()
                   && (declaration.Accessibility == Accessibility.Global
                       || declaration.Accessibility == Accessibility.Public
                       || (isSameProject && declaration.Accessibility == Accessibility.Friend));
        }

        private bool IsGlobalProcedure(Declaration declaration)
        {
            // a procedure is global if it's a Sub or Function (properties are never global).
            // since we have no visibility on module attributes,
            // we must assume a class member can be called from a default instance.

            if (declaration.DeclarationType != DeclarationType.Procedure ||
                declaration.DeclarationType != DeclarationType.Function)
            {
                return false;
            }

            // Friend members are only visible within the same project.
            var isSameProject = declaration.Project == _qualifiedName.Project;

            // implicit (unspecified) accessibility makes a member Public,
            // so if it's in the same project, it's public whenever it's not explicitly Private:
            return isSameProject && declaration.Accessibility == Accessibility.Friend
                   || declaration.Accessibility != Accessibility.Private;
        }

        private bool IsDeclarativeContext(VBAParser.AmbiguousIdentifierContext context)
        {
            return IsDeclarativeParentContext(context.Parent);
        }

        private bool IsDeclarativeContext(VBAParser.CertainIdentifierContext context)
        {
            return IsDeclarativeParentContext(context.Parent);
        }

        private static readonly Type[] DeclarativeContextTypes =
        {
            typeof (VBAParser.VariableSubStmtContext),
            typeof (VBAParser.ConstSubStmtContext),
            typeof (VBAParser.ArgContext),
            typeof (VBAParser.SubStmtContext),
            typeof (VBAParser.FunctionStmtContext),
            typeof (VBAParser.PropertyGetStmtContext),
            typeof (VBAParser.PropertyLetStmtContext),
            typeof (VBAParser.PropertySetStmtContext),
            typeof (VBAParser.TypeStmtContext),
            typeof (VBAParser.TypeStmt_ElementContext),
            typeof (VBAParser.EnumerationStmtContext),
            typeof (VBAParser.EnumerationStmt_ConstantContext),
            typeof (VBAParser.DeclareStmtContext),
            typeof (VBAParser.EventStmtContext)
        };

        private bool IsDeclarativeParentContext(RuleContext parentContext)
        {
            return DeclarativeContextTypes.Contains(parentContext.GetType());
        }
    }
}