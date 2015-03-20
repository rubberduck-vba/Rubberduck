using System;

namespace Rubberduck.Parsing.Symbols
{
    public class DeclarationSymbolsListener : VBABaseListener
    {
        private readonly Declarations _declarations = new Declarations();
        public Declarations Declarations { get { return _declarations; } }

        private readonly int _projectHashCode;
        private readonly string _projectName;
        private readonly string _componentName;

        private string _currentScope;

        public DeclarationSymbolsListener(int projectHashCode, string projectName, string componentName, Accessibility componentAccessibility, DeclarationType declarationType)
        {
            _projectHashCode = projectHashCode;
            _projectName = projectName;
            _componentName = componentName;

            SetCurrentScope();
            _declarations.Add(new Declaration(_projectHashCode, _projectName, _projectName, _componentName, _componentName, componentAccessibility, declarationType));
        }

        private Declaration CreateDeclaration(string identifierName, Accessibility accessibility, DeclarationType declarationType)
        {
            return new Declaration(_projectHashCode, _currentScope, _projectName, _componentName, identifierName, accessibility, declarationType);
        }

        /// <summary>
        /// Gets the <c>Accessibility</c> for a procedure member.
        /// </summary>
        /// <param name="visibilityContext"></param>
        /// <returns>Returns <c>Public</c> by default.</returns>
        private Accessibility GetProcedureAccessibility(VBAParser.VisibilityContext visibilityContext)
        {
            var visibility = visibilityContext == null
                ? "Public"
                : visibilityContext.GetText();

            return (Accessibility)Enum.Parse(typeof(Accessibility), visibility);
        }

        /// <summary>
        /// Gets the <c>Accessibility</c> for a non-procedure member.
        /// </summary>
        /// <param name="visibilityContext"></param>
        /// <returns>Returns <c>Private</c> by default.</returns>
        private Accessibility GetMemberAccessibility(VBAParser.VisibilityContext visibilityContext)
        {
            var visibility = visibilityContext == null
                ? "Private"
                : visibilityContext.GetText();

            return (Accessibility)Enum.Parse(typeof(Accessibility), visibility);
        }

        /// <summary>
        /// Sets current scope to module-level.
        /// </summary>
        private void SetCurrentScope()
        {
            _currentScope = _projectName + "." + _componentName;
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
            var accessibility = GetProcedureAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, accessibility, DeclarationType.Procedure));
            SetCurrentScope(name);
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, accessibility, DeclarationType.Function));
            SetCurrentScope(name);
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, accessibility, DeclarationType.PropertyGet));
            SetCurrentScope(name);
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, accessibility, DeclarationType.PropertyLet));
            SetCurrentScope(name);
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, accessibility, DeclarationType.PropertySet));
            SetCurrentScope(name);
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterEventStmt(VBAParser.EventStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, accessibility, DeclarationType.Event));
            SetCurrentScope(name);
        }

        public override void ExitEventStmt(VBAParser.EventStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterDeclareStmt(VBAParser.DeclareStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, accessibility, DeclarationType.LibraryFunction));
            SetCurrentScope(name);
        }

        public override void ExitDeclareStmt(VBAParser.DeclareStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterArgList(VBAParser.ArgListContext context)
        {
            var args = context.arg();
            foreach (var argContext in args)
            {
                _declarations.Add(CreateDeclaration(argContext.ambiguousIdentifier().GetText(), Accessibility.Implicit, DeclarationType.Parameter));
            }
        }

        public override void EnterVariableSubStmt(VBAParser.VariableSubStmtContext context)
        {
            var parent = (VBAParser.VariableStmtContext)context.Parent;
            var accessibility = GetMemberAccessibility(parent.visibility());
            
            _declarations.Add(CreateDeclaration(context.ambiguousIdentifier().GetText(), accessibility, DeclarationType.Variable));
        }

        public override void EnterConstSubStmt(VBAParser.ConstSubStmtContext context)
        {
            var parent = (VBAParser.ConstStmtContext)context.Parent;
            var accessibility = GetMemberAccessibility(parent.visibility());

            _declarations.Add(CreateDeclaration(context.ambiguousIdentifier().GetText(), accessibility, DeclarationType.Constant));
        }

        public override void EnterTypeStmt(VBAParser.TypeStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, accessibility, DeclarationType.UserDefinedType));
            SetCurrentScope(name);
        }

        public override void ExitTypeStmt(VBAParser.TypeStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterTypeStmt_Element(VBAParser.TypeStmt_ElementContext context)
        {
            _declarations.Add(CreateDeclaration(context.ambiguousIdentifier().GetText(), Accessibility.Implicit, DeclarationType.UserDefinedTypeMember));
        }

        public override void EnterEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var name = context.ambiguousIdentifier().GetText();

            _declarations.Add(CreateDeclaration(name, accessibility, DeclarationType.Enumeration));
            SetCurrentScope(name);
        }

        public override void ExitEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterEnumerationStmt_Constant(VBAParser.EnumerationStmt_ConstantContext context)
        {
            _declarations.Add(CreateDeclaration(context.ambiguousIdentifier().GetText(), Accessibility.Implicit, DeclarationType.EnumerationMember));
        }
    }
}
