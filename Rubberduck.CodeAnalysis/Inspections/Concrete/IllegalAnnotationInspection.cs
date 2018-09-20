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
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class IllegalAnnotationInspection : ParseTreeInspectionBase
    {
        public IllegalAnnotationInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new IllegalAttributeAnnotationsListener(state);
        }
        
        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts.Select(context => 
                new QualifiedContextInspectionResult(this, 
                string.Format(InspectionResults.IllegalAnnotationInspection, ((VBAParser.AnnotationContext)context.Context).annotationName().GetText()), context));
        }

        public class IllegalAttributeAnnotationsListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly RubberduckParserState _state;

            private Lazy<Declaration> _module;
            private Lazy<IDictionary<string, Declaration>> _members;

            public IllegalAttributeAnnotationsListener(RubberduckParserState state)
            {
                _state = state;
            }

            private readonly List<QualifiedContext<ParserRuleContext>> _contexts =
                new List<QualifiedContext<ParserRuleContext>>();

            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            private bool _isFirstMemberProcessed;

            public void ClearContexts()
            {
                _contexts.Clear();
                _isFirstMemberProcessed = false;
            }

            #region scoping
            private Declaration _currentScopeDeclaration;
            private bool _hasMembers;

            private void SetCurrentScope(string memberName = null)
            {
                _hasMembers = !string.IsNullOrEmpty(memberName);
                // this is a one-time toggle until contexts are reset
                _isFirstMemberProcessed |= _hasMembers; 
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

            public override void EnterModule(VBAParser.ModuleContext context)
            {
                _module = new Lazy<Declaration>(() => _state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .SingleOrDefault(m => m.QualifiedName.QualifiedModuleName.Equals(CurrentModuleName)));

                _members = new Lazy<IDictionary<string, Declaration>>(() => _state.DeclarationFinder
                    .Members(CurrentModuleName)
                    .GroupBy(m => m.IdentifierName)
                    .ToDictionary(m => m.Key, m => m.First()));

                // we did not process the first member of the module we just entered, so reset here
                _isFirstMemberProcessed = false; 
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

            public override void EnterSubStmt(VBAParser.SubStmtContext context)
            {
                SetCurrentScope(Identifier.GetName(context.subroutineName()));
            }

            public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
            {
                SetCurrentScope(Identifier.GetName(context.functionName()));
            }

            public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
            {
                SetCurrentScope(Identifier.GetName(context.functionName()));
            }

            public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
            {
                SetCurrentScope(Identifier.GetName(context.subroutineName()));
            }

            public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
            {
                SetCurrentScope(Identifier.GetName(context.subroutineName()));
            }
            #endregion

            public override void ExitAnnotation(VBAParser.AnnotationContext context)
            {
                var name = Identifier.GetName(context.annotationName().unrestrictedIdentifier());
                var annotationType = (AnnotationType) Enum.Parse(typeof (AnnotationType), name, true);

                var moduleHasMembers = _members.Value.Any();

                var isMemberAnnotation = annotationType.HasFlag(AnnotationType.MemberAnnotation);
                var isModuleAnnotation = annotationType.HasFlag(AnnotationType.ModuleAnnotation);

                var isModuleAnnotatedForMemberAnnotation = isMemberAnnotation
                    && (_currentScopeDeclaration?.DeclarationType.HasFlag(DeclarationType.Module) ?? false);

                var isMemberAnnotatedForModuleAnnotation = isModuleAnnotation 
                    && (_currentScopeDeclaration?.DeclarationType.HasFlag(DeclarationType.Member) ?? false);

                var isIllegal = !(isMemberAnnotation && moduleHasMembers && !_isFirstMemberProcessed) &&
                                (isMemberAnnotatedForModuleAnnotation || isModuleAnnotatedForMemberAnnotation);

                if (isIllegal)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}