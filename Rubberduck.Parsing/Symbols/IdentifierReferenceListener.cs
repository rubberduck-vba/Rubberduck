using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReferenceListener : VBABaseListener
    {
        private readonly Declarations _declarations;
        private readonly QualifiedModuleName _qualifiedName;

        private string _currentScope;

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
        }

        /// <summary>
        /// Sets current scope to specified module member.
        /// </summary>
        /// <param name="name">The name of the member owning the current scope.</param>
        private void SetCurrentScope(string name)
        {
            _currentScope = _qualifiedName + "." + name;
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
            SetCurrentScope(context.ambiguousIdentifier().GetText());
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText());
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SetCurrentScope(context.ambiguousIdentifier().GetText());
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
            var call = Resolve(leftSide.iCS_S_ProcedureOrArrayCall(), out context)
                       ?? Resolve(leftSide.iCS_S_VariableOrProcedureCall(), out context)
                       ?? Resolve(leftSide.iCS_S_DictionaryCall(), out context)
                       ?? Resolve(leftSide.iCS_S_MembersCall(), out context);

            return context;
        }

        private VBAParser.AmbiguousIdentifierContext EnterDictionaryCall(VBAParser.DictionaryCallStmtContext dictionaryCall, VBAParser.AmbiguousIdentifierContext parentIdentifier = null)
        {
            if (dictionaryCall == null)
            {
                return null;
            }

            if (parentIdentifier != null)
            {
                if (!EnterIdentifier(parentIdentifier, parentIdentifier.GetSelection()))
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

            if (context.Parent is VBAParser.ForNextStmtContext 
                || context.Parent is VBAParser.ForEachStmtContext)
            {
                EnterIdentifier(context, selection, true);
            }
            else
            {
                EnterIdentifier(context, selection);
            }
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

                if (!declaration.References.Select(r => r.Context).Contains(reference.Context))
                {
                    declaration.AddReference(reference);
                    return true;
                }
                // note: non-matching names are not necessarily undeclared identifiers, e.g. "String" in "Dim foo As String".
            }

            return false;
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

        private Declaration Resolve(VBAParser.ICS_S_ProcedureOrArrayCallContext context, out VBAParser.AmbiguousIdentifierContext identifierContext)
        {
            if (context == null)
            {
                identifierContext = null;
                return null;
            }

            var identifier = context.ambiguousIdentifier();
            var name = identifier.GetText();

            var procedure = FindProcedureDeclaration(name, identifier);
            var result = procedure ?? FindVariableDeclaration(name, identifier);

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
            return Resolve(context, out discarded);
        }

        private Declaration Resolve(VBAParser.ICS_S_VariableOrProcedureCallContext context, out VBAParser.AmbiguousIdentifierContext identifierContext)
        {
            if (context == null)
            {
                identifierContext = null;
                return null;
            }

            var identifier = context.ambiguousIdentifier();
            var name = identifier.GetText();

            var procedure = FindProcedureDeclaration(name, identifier);
            var result = procedure ?? FindVariableDeclaration(name, identifier);

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
            return Resolve(context, out discarded);
        }

        private Declaration Resolve(VBAParser.ICS_S_DictionaryCallContext context, out VBAParser.AmbiguousIdentifierContext identifierContext, VBAParser.AmbiguousIdentifierContext parentIdentifier = null)
        {
            if (context == null)
            {
                identifierContext = null;
                return null;
            }

            var identifier = EnterDictionaryCall(context.dictionaryCallStmt(), parentIdentifier);
            var name = identifier.GetText();

            var result = FindVariableDeclaration(name, identifier);

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
            return Resolve(context, out discarded, parentIdentifier);
        }

        private Declaration Resolve(VBAParser.ICS_S_MembersCallContext context, out VBAParser.AmbiguousIdentifierContext identifierContext)
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
            return Resolve(context, out discarded);
        }

        private Declaration Resolve(VBAParser.ICS_B_ProcedureCallContext context)
        {
            var name = context.certainIdentifier().GetText();
            return FindProcedureDeclaration(name, context.certainIdentifier()); // note: is this a StackOverflowException waiting to bite me?
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
                //|| item.DeclarationType == DeclarationType.Module
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

        private Declaration FindProcedureDeclaration(string procedureName, ParserRuleContext context)
        {
            var matches = _declarations[procedureName]
                .Where(declaration => ProcedureDeclarations.Contains(declaration.DeclarationType))
                .Where(IsInScope);

            var procedure = GetClosestScopeDeclaration(matches, context);
            return procedure;
        }

        private Declaration FindVariableDeclaration(string procedureName, ParserRuleContext context)
        {
            var matches = _declarations[procedureName]
                .Where(declaration => declaration.DeclarationType == DeclarationType.Variable || declaration.DeclarationType == DeclarationType.Parameter)
                .Where(IsInScope);

            var variable = GetClosestScopeDeclaration(matches, context);
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

        private Declaration GetClosestScopeDeclaration(IEnumerable<Declaration> declarations, ParserRuleContext context, DeclarationType accessorType = DeclarationType.PropertyGet)
        {
            if (context.Parent.Parent.Parent is VBAParser.AsTypeClauseContext)
            {
                accessorType = DeclarationType.Class;
            }

            var matches = declarations as IList<Declaration> ?? declarations.ToList();
            var currentScope = matches.FirstOrDefault(declaration => 
                IsCurrentScopeMember(accessorType, declaration)
                && (declaration.DeclarationType == accessorType
                || accessorType == DeclarationType.PropertyGet));

            if (currentScope != null)
            {
                //return currentScope;
            }

            var moduleScope = matches.SingleOrDefault(declaration => declaration.Scope == ModuleScope);
            if (moduleScope != null)
            {
                return moduleScope;
            }

            if (matches.Count == 1)
            {
                return matches[0];
            }

            var memberProcedureCallContext = context.Parent as VBAParser.ICS_B_MemberProcedureCallContext;
            if (memberProcedureCallContext != null)
            {
                return Resolve(memberProcedureCallContext);
                var parent = memberProcedureCallContext;
                var parentMemberName = memberProcedureCallContext.ambiguousIdentifier().GetText(); 
                var matchingParents = _declarations.Items.Where(d => d.IdentifierName == parentMemberName 
                    && (d.DeclarationType == DeclarationType.Class || d.DeclarationType == DeclarationType.UserDefinedType));

                var parentType = _withQualifiers.Any() 
                    ? _withQualifiers.Peek() 
                    : matches.SingleOrDefault(m => 
                        matchingParents.Any(p => 
                            (p.DeclarationType == DeclarationType.Class && m.ComponentName == p.AsTypeName)
                            || (p.DeclarationType == DeclarationType.UserDefinedType)));

                return parentType == null ? null : matches.SingleOrDefault(m => m.ParentScope == parentType.Scope);
            }

            return matches.SingleOrDefault(m => m.ParentScope == _currentScope);
        }

        private bool IsCurrentScopeMember(DeclarationType accessorType, Declaration declaration)
        {
            if (declaration.Scope != _currentScope && accessorType != DeclarationType.Class)
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