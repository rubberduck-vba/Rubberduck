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
    public sealed class MissingAnnotationInspection : InspectionBase, IParseTreeInspection
    {
        public MissingAnnotationInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
            Listener = new MissingMemberAttributeListener(state.DeclarationFinder);
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;
        public IInspectionListener Listener { get; }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts.Select(context => new MissingAnnotationInspectionResult(this, context, context.MemberName));
        }

        public class MissingMemberAttributeListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly DeclarationFinder _finder;
            private readonly HashSet<string> _attributeNames;

            public MissingMemberAttributeListener(DeclarationFinder finder)
            {
                _finder = finder;

                _attributeNames = new HashSet<string>(AnnotationType.Attribute.GetType()
                    .GetFields()
                    .Where(field => field.GetCustomAttributes(typeof(AttributeAnnotationAttribute), true).Any())
                    .SelectMany(a => a.GetCustomAttributes(typeof(AttributeAnnotationAttribute), true)
                        .Cast<AttributeAnnotationAttribute>())
                    .Select(a => a.AttributeName));
            }

            private readonly List<QualifiedContext<ParserRuleContext>> _contexts =
                new List<QualifiedContext<ParserRuleContext>>();

            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            private Declaration _currentScope;

            private void SetCurrentScope(string name)
            {
                _currentScope = _finder
                    .Members(CurrentModuleName)
                    .Single(m => m.IdentifierName == name);
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

            public override void ExitAnnotation(VBAParser.AnnotationContext context)
            {
                base.ExitAnnotation(context);
            }

            public override void ExitAttributeStmt(VBAParser.AttributeStmtContext context)
            {
                if(_currentScope == null)
                {
                    // not scoped yet can't be a member attribute
                    return;
                }

                var name = context.attributeName().GetText();


                if(!_currentScope.Annotations.Any(a => a.AnnotationType.HasFlag(AnnotationType.Attribute)
                                                       &&
                                                       _attributeNames.Select(
                                                           n => $"{_currentScope.IdentifierName}.{n}")
                                                           .All(n => n != name)))
                {
                    // current scope is missing an annotation for this attribute
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }

                base.ExitAttributeStmt(context);
            }
        }
    }
}