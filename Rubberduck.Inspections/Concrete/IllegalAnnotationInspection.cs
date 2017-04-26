using System;
using System.Collections.Generic;
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
    public sealed class IllegalAnnotationInspection : InspectionBase, IParseTreeInspection
    {
        public IllegalAnnotationInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Warning)
        {
            Listener = new IllegalAttributeAnnotationsListener();
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;
        public IInspectionListener Listener { get; }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts.Select(context => new IllegalAnnotationInspectionResult(this, context, context.MemberName));
        }

        public class IllegalAttributeAnnotationsListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly IDictionary<AnnotationType, int> _annotationCounts;

            private static readonly AnnotationType[] AnnotationTypes = Enum.GetValues(typeof(AnnotationType)).Cast<AnnotationType>().ToArray(); 

            public IllegalAttributeAnnotationsListener()
            {
                _annotationCounts = AnnotationTypes.ToDictionary(a => a, a => 0);
            }

            private readonly List<QualifiedContext<ParserRuleContext>> _contexts =
                new List<QualifiedContext<ParserRuleContext>>();

            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            #region scoping
            private string _currentScope;

            private void SetCurrentScope(string name)
            {
                _currentScope = name;
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
                var annotationType = (AnnotationType) Enum.Parse(typeof (AnnotationType), name);
                _annotationCounts[annotationType]++;

                var isPerModule = annotationType.HasFlag(AnnotationType.ModuleAnnotation);
                var isMemberOnModule = _currentScope != null && isPerModule;

                var isPerMember = annotationType.HasFlag(AnnotationType.MemberAnnotation);
                var isModuleOnMember = _currentScope == null && isPerMember;

                var isOnlyAllowedOnce = isPerModule || isPerMember;

                if (isOnlyAllowedOnce && _annotationCounts[annotationType] > 1 || isModuleOnMember || isMemberOnModule)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}