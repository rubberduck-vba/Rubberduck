using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class DeclarationSymbolsListener : VBABaseListener
    {
        private readonly QualifiedModuleName _qualifiedName;
        private readonly Declaration _moduleDeclaration;
        private readonly Declaration _projectDeclaration;

        private string _currentScope;
        private Declaration _parentDeclaration;

        private readonly IEnumerable<CommentNode> _comments;
        private readonly CancellationToken _token;

        public DeclarationSymbolsListener(QualifiedModuleName qualifiedName, Accessibility componentAccessibility, vbext_ComponentType type, IEnumerable<CommentNode> comments, CancellationToken token)
        {
            _qualifiedName = qualifiedName;
            _comments = comments;
            _token = token;

            var declarationType = type == vbext_ComponentType.vbext_ct_StdModule
                ? DeclarationType.Module
                //: result.Component.Type == vbext_ComponentType.vbext_ct_MSForm 
                //    ? DeclarationType.UserForm
                //    : result.Component.Type == vbext_ComponentType.vbext_ct_Document
                //        ? DeclarationType.Document
                : DeclarationType.Class;

            var project = _qualifiedName.Component.Collection.Parent;
            var projectQualifiedName = new QualifiedModuleName(project);

            _projectDeclaration = new Declaration(
                projectQualifiedName.QualifyMemberName(project.Name),
                null, null, project.Name, false, false, Accessibility.Implicit, DeclarationType.Project, false);

            _moduleDeclaration = new Declaration(
                _qualifiedName.QualifyMemberName(_qualifiedName.Component.Name),
                _projectDeclaration,
                _qualifiedName.Project.Name,
                _qualifiedName.Component.Name,
                false, 
                false,
                componentAccessibility,
                declarationType,
                null,
                Selection.Home,
                false, 
                FindAnnotations());

            SetCurrentScope();
        }

        private string FindAnnotations()
        {
            if (_comments == null)
            {
                return null;
            }

            var lastDeclarationsSectionLine = _qualifiedName.Component.CodeModule.CountOfDeclarationLines;
            var annotations = _comments.Where(comment => comment.QualifiedSelection.QualifiedName.Equals(_qualifiedName)
                && comment.QualifiedSelection.Selection.EndLine <= lastDeclarationsSectionLine
                && comment.CommentText.StartsWith("@")).ToArray();

            if (annotations.Any())
            {
                return string.Join("\n", annotations.Select(annotation => annotation.CommentText));
            }

            return null;
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

        public void CreateModuleDeclarations()
        {
            OnNewDeclaration(_projectDeclaration);
            OnNewDeclaration(_moduleDeclaration);

            var component = _moduleDeclaration.QualifiedName.QualifiedModuleName.Component;
            if (component.Type == vbext_ComponentType.vbext_ct_MSForm || component.Designer != null)
            {
                DeclareControlsAsMembers(component);
            }
        }

        public event EventHandler<DeclarationEventArgs> NewDeclaration;
        private void OnNewDeclaration(Declaration declaration)
        {
            _token.ThrowIfCancellationRequested();

            var handler = NewDeclaration;
            if (handler != null)
            {
                handler.Invoke(this, new DeclarationEventArgs(declaration));
            }
        }

        /// <summary>
        /// Scans form designer to create a public, self-assigned field for each control on a form.
        /// </summary>
        /// <remarks>
        /// These declarations are meant to be used to identify control event procedures.
        /// </remarks>
        private void DeclareControlsAsMembers(VBComponent form)
        {
            var designer = form.Designer;
            if (designer == null)
            {
                return;
            }

            // using dynamic typing here, because not only MSForms could have a Controls collection (e.g. MS-Access forms are 'document' modules).
            foreach (var control in ((dynamic)designer).Controls)
            {
                var declaration = new Declaration(_qualifiedName.QualifyMemberName(control.Name), _parentDeclaration, _currentScope, "Control", true, true, Accessibility.Public, DeclarationType.Control, null, Selection.Home);
                OnNewDeclaration(declaration);
            }
        }

        private Declaration CreateDeclaration(string identifierName, string asTypeName, Accessibility accessibility, DeclarationType declarationType, ParserRuleContext context, Selection selection, bool selfAssigned = false, bool withEvents = false)
        {
            var annotations = FindAnnotations(selection.StartLine);
            var result = new Declaration(new QualifiedMemberName(_qualifiedName, identifierName), _parentDeclaration, _currentScope, asTypeName, selfAssigned, withEvents, accessibility, declarationType, context, selection, false, annotations);

            OnNewDeclaration(result);
            return result;
        }

        /// <summary>
        /// Gets the <c>Accessibility</c> for a procedure member.
        /// </summary>
        private Accessibility GetProcedureAccessibility(VBAParser.VisibilityContext visibilityContext)
        {
            var visibility = visibilityContext == null
                ? "Implicit" // "Public"
                : visibilityContext.GetText();

            return (Accessibility)Enum.Parse(typeof(Accessibility), visibility);
        }

        /// <summary>
        /// Gets the <c>Accessibility</c> for a non-procedure member.
        /// </summary>
        private Accessibility GetMemberAccessibility(VBAParser.VisibilityContext visibilityContext)
        {
            var visibility = visibilityContext == null
                ? "Implicit" // "Private"
                : visibilityContext.GetText();

            return (Accessibility)Enum.Parse(typeof(Accessibility), visibility, true);
        }

        /// <summary>
        /// Sets current scope to module-level.
        /// </summary>
        private void SetCurrentScope()
        {
            _currentScope = _qualifiedName.ToString();
            _parentDeclaration = _moduleDeclaration;
        }

        /// <summary>
        /// Sets current scope to specified module member.
        /// </summary>
        /// <param name="procedureDeclaration"></param>
        /// <param name="name">The name of the member owning the current scope.</param>
        private void SetCurrentScope(Declaration procedureDeclaration, string name)
        {
            _currentScope = _qualifiedName + "." + name;
        }

        public override void EnterOptionBaseStmt(VBAParser.OptionBaseStmtContext context)
        {
            OnNewDeclaration(CreateDeclaration(context.GetText(), string.Empty, Accessibility.Implicit, DeclarationType.ModuleOption, context, context.GetSelection()));
        }

        public override void EnterOptionCompareStmt(VBAParser.OptionCompareStmtContext context)
        {
            OnNewDeclaration(CreateDeclaration(context.GetText(), string.Empty, Accessibility.Implicit, DeclarationType.ModuleOption, context, context.GetSelection()));
        }

        public override void EnterOptionExplicitStmt(VBAParser.OptionExplicitStmtContext context)
        {
            OnNewDeclaration(CreateDeclaration(context.GetText(), string.Empty, Accessibility.Implicit, DeclarationType.ModuleOption, context, context.GetSelection()));
        }

        public override void ExitOptionPrivateModuleStmt(VBAParser.OptionPrivateModuleStmtContext context)
        {
            OnNewDeclaration(CreateDeclaration(context.GetText(), string.Empty, Accessibility.Implicit, DeclarationType.ModuleOption, context, context.GetSelection()));
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var identifier = context.ambiguousIdentifier();
            if (identifier == null)
            {
                return;
            }

            var name = context.ambiguousIdentifier().GetText();
            var declaration = CreateDeclaration(name, null, accessibility, DeclarationType.Procedure, context, context.ambiguousIdentifier().GetSelection());
            OnNewDeclaration(declaration);
            SetCurrentScope(declaration, name);
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var identifier = context.ambiguousIdentifier();
            if (identifier == null)
            {
                return;
            }
            var name = identifier.GetText();

            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null 
                ? Tokens.Variant 
                : asTypeClause.type().GetText();

            var declaration = CreateDeclaration(name, asTypeName, accessibility, DeclarationType.Function, context, context.ambiguousIdentifier().GetSelection());
            OnNewDeclaration(declaration);
            SetCurrentScope(declaration, name);
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var identifier = context.ambiguousIdentifier();
            if (identifier == null)
            {
                return;
            }
            var name = identifier.GetText();

            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();

            var declaration = CreateDeclaration(name, asTypeName, accessibility, DeclarationType.PropertyGet, context, context.ambiguousIdentifier().GetSelection());

            OnNewDeclaration(declaration);
            SetCurrentScope(declaration, name);
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var identifier = context.ambiguousIdentifier();
            if (identifier == null)
            {
                return;
            }
            var name = identifier.GetText();

            var declaration = CreateDeclaration(name, null, accessibility, DeclarationType.PropertyLet, context, context.ambiguousIdentifier().GetSelection());
            OnNewDeclaration(declaration);
            SetCurrentScope(declaration, name);
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            var accessibility = GetProcedureAccessibility(context.visibility());
            var identifier = context.ambiguousIdentifier();
            if (identifier == null)
            {
                return;
            }
            var name = identifier.GetText();

            var declaration = CreateDeclaration(name, null, accessibility, DeclarationType.PropertySet, context, context.ambiguousIdentifier().GetSelection());

            OnNewDeclaration(declaration);
            SetCurrentScope(declaration, name);
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterEventStmt(VBAParser.EventStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var identifier = context.ambiguousIdentifier();
            if (identifier == null)
            {
                return;
            }
            var name = identifier.GetText();

            var declaration = CreateDeclaration(name, null, accessibility, DeclarationType.Event, context, context.ambiguousIdentifier().GetSelection());

            OnNewDeclaration(declaration);
            SetCurrentScope(declaration, name);
        }

        public override void ExitEventStmt(VBAParser.EventStmtContext context)
        {
            SetCurrentScope();
        }

        public override void EnterDeclareStmt(VBAParser.DeclareStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var nameContext = context.ambiguousIdentifier();
            if (nameContext == null)
            {
                return;
            }
            var name = nameContext.GetText();

            var hasReturnType = context.FUNCTION() != null;

            var asTypeClause = context.asTypeClause();
            var asTypeName = hasReturnType 
                                ? asTypeClause == null
                                    ? Tokens.Variant
                                    : asTypeClause.type().GetText() 
                                : null;

            var selection = nameContext.GetSelection();

            var declarationType = hasReturnType
                ? DeclarationType.LibraryFunction
                : DeclarationType.LibraryProcedure;

            var declaration = CreateDeclaration(name, asTypeName, accessibility, declarationType, context, selection);

            OnNewDeclaration(declaration);
            SetCurrentScope(declaration, name); // treat like a procedure block, to correctly scope parameters.
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
                var asTypeClause = argContext.asTypeClause();
                var asTypeName = asTypeClause == null
                    ? Tokens.Variant
                    : asTypeClause.type().GetText();

                var identifier = argContext.ambiguousIdentifier();
                if (identifier == null)
                {
                    return;
                }
                OnNewDeclaration(CreateDeclaration(identifier.GetText(), asTypeName, Accessibility.Implicit, DeclarationType.Parameter, argContext, identifier.GetSelection()));
            }
        }

        public override void EnterLineLabel(VBAParser.LineLabelContext context)
        {
            OnNewDeclaration(CreateDeclaration(context.ambiguousIdentifier().GetText(), null, Accessibility.Private, DeclarationType.LineLabel, context, context.ambiguousIdentifier().GetSelection(), true));
        }

        public override void EnterVariableSubStmt(VBAParser.VariableSubStmtContext context)
        {
            var parent = (VBAParser.VariableStmtContext)context.Parent.Parent;
            var accessibility = GetMemberAccessibility(parent.visibility());
            var identifier = context.ambiguousIdentifier();
            if (identifier == null)
            {
                return;
            }
            var name = identifier.GetText();

            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();

            var withEvents = parent.WITHEVENTS() != null;
            var selfAssigned = asTypeClause != null && asTypeClause.NEW() != null;
            
            OnNewDeclaration(CreateDeclaration(name, asTypeName, accessibility, DeclarationType.Variable, context, context.ambiguousIdentifier().GetSelection(), selfAssigned, withEvents));
        }

        public override void EnterConstSubStmt(VBAParser.ConstSubStmtContext context)
        {
            var parent = (VBAParser.ConstStmtContext)context.Parent;
            var accessibility = GetMemberAccessibility(parent.visibility());

            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();

            var identifier = context.ambiguousIdentifier();
            if (identifier == null)
            {
                return;
            }
            var name = identifier.GetText();
            var value = context.valueStmt().GetText();
            var declaration = new ValuedDeclaration(new QualifiedMemberName(_qualifiedName, name), _parentDeclaration, _currentScope, asTypeName, accessibility, DeclarationType.Constant, value, context, identifier.GetSelection());

            OnNewDeclaration(declaration);
        }

        public override void EnterTypeStmt(VBAParser.TypeStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var identifier = context.ambiguousIdentifier();
            if (identifier == null)
            {
                return;
            }
            var name = identifier.GetText();

            var declaration = CreateDeclaration(name, null, accessibility, DeclarationType.UserDefinedType, context, context.ambiguousIdentifier().GetSelection());

            OnNewDeclaration(declaration);
            _parentDeclaration = declaration; // treat members as child declarations, but keep them scoped to module
        }

        public override void ExitTypeStmt(VBAParser.TypeStmtContext context)
        {
            _parentDeclaration = _moduleDeclaration;
        }

        public override void EnterTypeStmt_Element(VBAParser.TypeStmt_ElementContext context)
        {
            var asTypeClause = context.asTypeClause();
            var asTypeName = asTypeClause == null
                ? Tokens.Variant
                : asTypeClause.type().GetText();

            OnNewDeclaration(CreateDeclaration(context.ambiguousIdentifier().GetText(), asTypeName, Accessibility.Implicit, DeclarationType.UserDefinedTypeMember, context, context.ambiguousIdentifier().GetSelection()));
        }

        public override void EnterEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            var accessibility = GetMemberAccessibility(context.visibility());
            var identifier = context.ambiguousIdentifier();
            if (identifier == null)
            {
                return;
            }
            var name = identifier.GetText();

            var declaration = CreateDeclaration(name, null, accessibility, DeclarationType.Enumeration, context, context.ambiguousIdentifier().GetSelection());

            OnNewDeclaration(declaration);
            _parentDeclaration = declaration; // treat members as child declarations, but keep them scoped to module
        }

        public override void ExitEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            _parentDeclaration = _moduleDeclaration;
        }

        public override void EnterEnumerationStmt_Constant(VBAParser.EnumerationStmt_ConstantContext context)
        {
            OnNewDeclaration(CreateDeclaration(context.ambiguousIdentifier().GetText(), null, Accessibility.Implicit, DeclarationType.EnumerationMember, context, context.ambiguousIdentifier().GetSelection()));
        }
    }
}
