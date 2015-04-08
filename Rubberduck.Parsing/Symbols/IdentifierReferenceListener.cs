using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
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

        private string ModuleScope { get { return _qualifiedName.ProjectName + "." + _qualifiedName.ModuleName; } }

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
            _currentScope = _qualifiedName.ProjectName + "." + _qualifiedName.ModuleName + "." + name;
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
            var target = FindAssignmentTarget(leftSide);
            if (target != null)
            {
                EnterIdentifier(target, target.GetSelection(), true);
            }
        }

        public override void EnterSetStmt(VBAParser.SetStmtContext context)
        {
            var leftSide = context.implicitCallStmt_InStmt();
            var target = FindAssignmentTarget(leftSide);
            if (target != null)
            {
                EnterIdentifier(target, target.GetSelection(), true);
            }
        }

        private VBAParser.AmbiguousIdentifierContext FindAssignmentTarget(VBAParser.ImplicitCallStmt_InStmtContext leftSide)
        {
            // todo: refactor!!

            var procedureOrArrayCall = leftSide.iCS_S_ProcedureOrArrayCall();
            if (procedureOrArrayCall != null)
            {
                return EnterDictionaryCall(procedureOrArrayCall.dictionaryCallStmt(),
                    procedureOrArrayCall.ambiguousIdentifier())
                                   ?? procedureOrArrayCall.ambiguousIdentifier();
            }

            var variableOrProcedureCall = leftSide.iCS_S_VariableOrProcedureCall();
            if (variableOrProcedureCall != null)
            {
                return EnterDictionaryCall(variableOrProcedureCall.dictionaryCallStmt(),
                    variableOrProcedureCall.ambiguousIdentifier())
                                   ?? variableOrProcedureCall.ambiguousIdentifier();
            }

            var dictionaryCall = leftSide.iCS_S_DictionaryCall();
            if (dictionaryCall != null && dictionaryCall.dictionaryCallStmt() != null)
            {
                return EnterDictionaryCall(dictionaryCall.dictionaryCallStmt());
            }

            var membersCall = leftSide.iCS_S_MembersCall();
            if (membersCall != null)
            {
                var members = membersCall.iCS_S_MemberCall();
                for (var index = 0; index < members.Count; index++)
                {
                    var member = members[index];
                    if (index < members.Count - 1)
                    {
                        var procOrArrayCall = member.iCS_S_ProcedureOrArrayCall();
                        if (procOrArrayCall != null)
                        {
                            var reference = EnterDictionaryCall(procOrArrayCall.dictionaryCallStmt(), procOrArrayCall.ambiguousIdentifier())
                                         ?? procOrArrayCall.ambiguousIdentifier();

                            if (reference != null)
                            {
                                EnterIdentifier(reference, reference.GetSelection());
                            }
                        }

                        var varOrProcCall = member.iCS_S_VariableOrProcedureCall();
                        if (varOrProcCall != null)
                        {
                            var reference = EnterDictionaryCall(varOrProcCall.dictionaryCallStmt(), varOrProcCall.ambiguousIdentifier())
                                         ?? varOrProcCall.ambiguousIdentifier();

                            if (reference != null)
                            {
                                EnterIdentifier(reference, reference.GetSelection());
                            }
                        }
                    }
                    else
                    {
                        var procOrArrayCall = member.iCS_S_ProcedureOrArrayCall();
                        if (procOrArrayCall != null)
                        {
                            return EnterDictionaryCall(procOrArrayCall.dictionaryCallStmt(), procOrArrayCall.ambiguousIdentifier())
                                ?? procOrArrayCall.ambiguousIdentifier();
                        }

                        var varOrProcCall = member.iCS_S_VariableOrProcedureCall();
                        if (varOrProcCall != null)
                        {
                            return EnterDictionaryCall(varOrProcCall.dictionaryCallStmt(), varOrProcCall.ambiguousIdentifier())
                                ?? varOrProcCall.ambiguousIdentifier();
                        }
                    }
                }
            }

            return null; // not possible unless grammar is modified.
        }

        private VBAParser.AmbiguousIdentifierContext EnterDictionaryCall(VBAParser.DictionaryCallStmtContext dictionaryCall, VBAParser.AmbiguousIdentifierContext parentIdentifier = null)
        {
            if (dictionaryCall == null)
            {
                return null;
            }

            if (parentIdentifier != null)
            {
                EnterIdentifier(parentIdentifier, parentIdentifier.GetSelection()); // we're referencing "member" in "member!field"
            }
            
            return dictionaryCall.ambiguousIdentifier();
        }

        public override void EnterAmbiguousIdentifier(VBAParser.AmbiguousIdentifierContext context)
        {
            if (IsDeclarativeContext(context))
            {
                return;
            }

            var selection = context.GetSelection();

            if (context.Parent is VBAParser.ForNextStmtContext)
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

        private void EnterIdentifier(ParserRuleContext context, Selection selection, bool isAssignmentTarget = false)
        {
            var name = context.GetText();
            var matches = _declarations[name].Where(IsInScope);

            var declaration = GetClosestScope(matches);
            if (declaration != null)
            {
                var reference = new IdentifierReference(_qualifiedName, name, selection, context, declaration, isAssignmentTarget);

                if (!declaration.References.Select(r => r.Context).Contains(reference.Context))
                {
                    declaration.AddReference(reference);
                }
                // note: non-matching names are not necessarily undeclared identifiers, e.g. "String" in "Dim foo As String".
            }
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
            if (callStatementA != null)
            {
                procedureName = callStatementA.ambiguousIdentifier().GetText();
            }
            else if(callStatementB != null)
            {
                procedureName = callStatementB.ambiguousIdentifier().GetText();
            }
            else if (callStatementC != null)
            {
                procedureName = callStatementC.ambiguousIdentifier().GetText();
            }
            else if (callStatementD != null)
            {
                procedureName = callStatementD.certainIdentifier().GetText();
            }

            var procedure = FindProcedureDeclaration(procedureName);
            if (procedure == null)
            {
                return;
            }

            var procScope = procedure.ParentScope + "." + procedure.IdentifierName;

            var arg = _declarations.Items.FirstOrDefault(declaration =>
                        declaration.ParentScope == procScope 
                        && declaration.DeclarationType == DeclarationType.Parameter);

            if (arg != null)
            {
                var reference = new IdentifierReference(_qualifiedName, arg.IdentifierName, context.GetSelection(), context, arg);
                arg.AddReference(reference);
            }
        }

        private Declaration FindProcedureDeclaration(string procedureName)
        {
            var matches = _declarations[procedureName]
                .Where(declaration => ProcedureDeclarations.Contains(declaration.DeclarationType))
                .Where(IsInScope);

            var procedure = GetClosestScope(matches);
            return procedure;
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
            if (declaration.DeclarationType == DeclarationType.Project)
            {
                return true; // a project name is always in scope anywhere
            }

            if (declaration.DeclarationType == DeclarationType.Module ||
                declaration.DeclarationType == DeclarationType.Class)
            {
                // todo: access component instancing properties to do this right (class)
                return true;
            }

            if (ProcedureDeclarations.Contains(declaration.DeclarationType))
            {
                if (declaration.Accessibility == Accessibility.Public)
                {
                    var result = declaration.Project.Equals(_qualifiedName.Project);
                    return result;
                }

                return declaration.QualifiedName.QualifiedModuleName == _qualifiedName;
            }

            return declaration.Scope == _currentScope
                   || declaration.Scope == ModuleScope
                   || IsGlobalField(declaration) 
                   || IsGlobalProcedure(declaration);
        }

        private Declaration GetClosestScope(IEnumerable<Declaration> declarations)
        {
            // this method (as does the rest of Rubberduck) assumes the VBA code is compilable.

            var matches = declarations as IList<Declaration> ?? declarations.ToList();
            var currentScope = matches.FirstOrDefault(declaration => declaration.Scope == _currentScope);
            if (currentScope != null)
            {
                return currentScope;
            }

            var moduleScope = matches.FirstOrDefault(declaration => declaration.Scope == ModuleScope);
            if (moduleScope != null)
            {
                return moduleScope;
            }

            // multiple matches in global scope??
            return matches.FirstOrDefault();
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

        private bool IsAssignmentContext(RuleContext context)
        {
            var parent = context.Parent;
            return parent.Parent.Parent is VBAParser.SetStmtContext // object reference assignment
                   || parent.Parent.Parent is VBAParser.LetStmtContext // value assignment
                   || parent is VBAParser.ForNextStmtContext // treat For loop as an assignment
                // todo: verify that we're not missing anything here (likely)
                ;
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