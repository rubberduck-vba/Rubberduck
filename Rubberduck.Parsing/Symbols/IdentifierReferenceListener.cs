using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReferenceListener : VBABaseListener
    {
        private readonly Declarations _declarations;

        private readonly int _projectHashCode;
        private readonly string _projectName;
        private readonly string _componentName;

        private string _currentScope;

        public IdentifierReferenceListener(VBComponentParseResult result, Declarations declarations)
            : this(result.QualifiedName.ProjectHashCode, result.QualifiedName.ProjectName, result.QualifiedName.ModuleName, declarations)
        { }

        public IdentifierReferenceListener(int projectHashCode, string projectName, string componentName, Declarations declarations)
        {
            _projectHashCode = projectHashCode;
            _projectName = projectName;
            _componentName = componentName;

            _declarations = declarations;

            SetCurrentScope();
        }

        private string ModuleScope { get { return _projectName + "." + _componentName; } }

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
            _currentScope = _projectName + "." + _componentName + "." + name;
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

        public override void EnterAmbiguousIdentifier(VBAParser.AmbiguousIdentifierContext context)
        {
            if (IsDeclarativeContext(context))
            {
                // skip declarations
                return;
            }

            var selection = context.GetSelection();
            EnterIdentifier(context, selection);
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

        private void EnterIdentifier(RuleContext context, Selection selection)
        {
            var name = context.GetText();
            var matches = _declarations[name].Where(IsInScope);

            var declaration = GetClosestScope(matches);
            if (declaration != null)
            {
                var isAssignment = IsAssignmentContext(context);
                var reference = new IdentifierReference(_projectName, _componentName, name, selection, isAssignment);

                declaration.AddReference(reference);
                // note: non-matching names are not necessarily undeclared identifiers, e.g. "String" in "Dim foo As String".
            }
        }

        private bool IsInScope(Declaration declaration)
        {
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
            var isSameProject = declaration.ProjectHashCode == _projectHashCode;

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
            var isSameProject = declaration.ProjectHashCode == _projectHashCode;

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

        private bool IsDeclarativeParentContext(RuleContext parentContext)
        {
            return parentContext is VBAParser.VariableSubStmtContext
                   || parentContext is VBAParser.ConstSubStmtContext
                   || parentContext is VBAParser.ArgContext
                   || parentContext is VBAParser.SubStmtContext
                   || parentContext is VBAParser.FunctionStmtContext
                   || parentContext is VBAParser.PropertyGetStmtContext
                   || parentContext is VBAParser.PropertyLetStmtContext
                   || parentContext is VBAParser.PropertySetStmtContext
                   || parentContext is VBAParser.TypeStmtContext
                   || parentContext is VBAParser.TypeStmt_ElementContext
                   || parentContext is VBAParser.EnumerationStmtContext
                   || parentContext is VBAParser.EnumerationStmt_ConstantContext
                   || parentContext is VBAParser.DeclareStmtContext
                   || parentContext is VBAParser.EventStmtContext;
        }
    }
}