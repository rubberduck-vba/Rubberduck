using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class MissingAttributeInspection : ParseTreeInspectionBase
    {
        public MissingAttributeInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new MissingMemberAttributeListener(state);
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.RubberduckOpportunities;
        public override IInspectionListener Listener { get; }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts.Select(context =>
            {
                var name = string.Format(InspectionsUI.MissingAttributeInspectionResultFormat, context.MemberName,
                    ((VBAParser.AnnotationContext) context.Context).annotationName().GetText());
                return new QualifiedContextInspectionResult(this, name, State, context);
            });
        }

        public class MissingMemberAttributeListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly RubberduckParserState _state;

            private readonly Lazy<Declaration> _module;
            private readonly Lazy<IDictionary<string, Declaration>> _members;

            public MissingMemberAttributeListener(RubberduckParserState state)
            {
                _state = state;
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
            private Declaration _currentScopeDeclaration;
            private bool _hasMembers;

            private void SetCurrentScope(IAnnotatedContext context, string memberName = null)
            {
                _hasMembers = !string.IsNullOrEmpty(memberName);
                _currentScopeDeclaration = _hasMembers ? _members.Value[memberName] : _module.Value;
            }

            public override void EnterModuleBody(VBAParser.ModuleBodyContext context)
            {
                _currentScopeDeclaration = _state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .Where(declaration => declaration.QualifiedName.QualifiedModuleName.Equals(CurrentModuleName))
                    .OrderBy(declaration => declaration.Selection)
                    .FirstOrDefault();
            }

            public override void ExitModule(VBAParser.ModuleContext context)
            {
                _currentScopeDeclaration = null;
            }

            public override void EnterModuleAttributes(VBAParser.ModuleAttributesContext context)
            {
                // note: using ModuleAttributesContext for module-scope

                if(_currentScopeDeclaration == null)
                {
                    // we're at the top of the module.
                    // everything we pick up between here and the module body, is module-scoped:
                    _currentScopeDeclaration = _state.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                        .SingleOrDefault(d => d.QualifiedName.QualifiedModuleName.Equals(CurrentModuleName));
                }
                else
                {
                    // DO NOT re-assign _currentScope here.
                    // we're at the end of the module and that attribute is actually scoped to the last procedure.
                    Debug.Assert(_currentScopeDeclaration != null); // deliberate no-op
                }
            }

            public override void ExitModuleDeclarations(VBAParser.ModuleDeclarationsContext context)
            {
                var firstMember = _members.Value.Values.OrderBy(d => d.Selection).FirstOrDefault();
                if (firstMember != null)
                {
                    _currentScopeDeclaration = firstMember;
                }
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

            public override void ExitAnnotation(VBAParser.AnnotationContext context)
            {
                var name = context.annotationName().GetText();
                if (_currentScopeDeclaration == null)
                {
                    // module-level annotation
                    var module = _state.DeclarationFinder.UserDeclarations(DeclarationType.Module).Single(m => m.QualifiedName.QualifiedModuleName.Equals(CurrentModuleName));
                    if (!module.Attributes.ContainsKey(name))
                    {
                        _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                    }
                }
                else
                {
                    // member-level annotation
                    var member = _members.Value.Single(m => m.Key.Equals(_currentScopeDeclaration.QualifiedName.MemberName));
                    if (!member.Value.Attributes.ContainsKey(name))
                    {
                        _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                    }
                }
            }
        }
    }
}