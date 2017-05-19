using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class MissingAnnotationInspection : ParseTreeInspectionBase
    {
        public MissingAnnotationInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            Listener = new MissingAnnotationListener(state);
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.RubberduckOpportunities;
        public override IInspectionListener Listener { get; }
        public override ParsePass Pass => ParsePass.AttributesPass;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts.Select(context =>
            {
                var member = string.IsNullOrEmpty(context.MemberName.MemberName)
                    ? context.ModuleName.Name
                    : context.MemberName.MemberName;

                var name = string.Format(InspectionsUI.MissingAnnotationInspectionResultFormat, 
                    member, ((VBAParser.AttributeStmtContext) context.Context).AnnotationType().ToString());

                return new QualifiedContextInspectionResult(this, name, State, context);
            });
        }

        public class MissingAnnotationListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly RubberduckParserState _state;

            private readonly Lazy<Declaration> _module;
            private readonly Lazy<IDictionary<string, Declaration>> _members;

            public MissingAnnotationListener(RubberduckParserState state)
            {
                _state = state;
                new VBAParserAnnotationFactory();

                _module = new Lazy<Declaration>(() => _state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .SingleOrDefault(m => m.QualifiedName.QualifiedModuleName.Equals(CurrentModuleName)));

                _members = new Lazy<IDictionary<string, Declaration>>(() => _state.DeclarationFinder
                    .Members(CurrentModuleName)
                    .GroupBy(m => m.IdentifierName)
                    .ToDictionary(m => m.Key, m => m.First()));
            }

            private readonly List<QualifiedContext<ParserRuleContext>> _contexts =
                new List<QualifiedContext<ParserRuleContext>>();

            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts() => _contexts.Clear();

            #region scoping
            private IAnnotatedContext _currentScope;
            private Declaration _currentScopeDeclaration;
            private bool _hasMembers;

            public override void EnterModuleBody(VBAParser.ModuleBodyContext context)
            {
                var firstMember = context.moduleBodyElement().FirstOrDefault()?.GetChild(0);
                _currentScope = firstMember as IAnnotatedContext;
                _currentScopeDeclaration = _state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .Where(declaration => declaration.QualifiedName.QualifiedModuleName.Equals(CurrentModuleName))
                    .OrderBy(declaration => declaration.Selection)
                    .FirstOrDefault();
            }

            public override void ExitModule(VBAParser.ModuleContext context)
            {
                _currentScope = null;
                _currentScopeDeclaration = null;
            }

            public override void EnterModuleAttributes(VBAParser.ModuleAttributesContext context)
            {
                // note: using ModuleAttributesContext for module-scope

                if (_currentScope == null)
                {
                    // we're at the top of the module.
                    // everything we pick up between here and the module body, is module-scoped:
                    _currentScope = context;
                    _currentScopeDeclaration = _state.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                        .SingleOrDefault(d => d.QualifiedName.QualifiedModuleName.Equals(CurrentModuleName));
                }
                else
                {
                    // DO NOT re-assign _currentScope here.
                    // we're at the end of the module and that attribute is actually scoped to the last procedure.
                    Debug.Assert(_currentScope != null); // deliberate no-op
                }
            }

            private void SetCurrentScope(IAnnotatedContext context, string memberName = null)
            {
                _hasMembers = !string.IsNullOrEmpty(memberName);
                _currentScope = context;
                _currentScopeDeclaration = _hasMembers ? _members.Value[memberName] : _module.Value;
            }

            public override void EnterSubStmt(VBAParser.SubStmtContext context)
            {
                SetCurrentScope(context, Identifier.GetName(context.subroutineName()));
            }

            public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
            {
                SetCurrentScope(context, Identifier.GetName(context.functionName()));
            }

            public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
            {
                SetCurrentScope(context, Identifier.GetName(context.functionName()));
            }

            public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
            {
                SetCurrentScope(context, Identifier.GetName(context.subroutineName()));
            }

            public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
            {
                SetCurrentScope(context, Identifier.GetName(context.subroutineName()));
            }
            #endregion
            
            public override void ExitAttributeStmt(VBAParser.AttributeStmtContext context)
            {
                Debug.Assert(_currentScopeDeclaration != null);
                var annotations = _currentScopeDeclaration.Annotations;

                var type = context.AnnotationType();
                if (type != null && annotations.All(a => a.AnnotationType != type))
                {
                    // attribute is mapped to an annotation, but current scope doesn't have that annotation:
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}