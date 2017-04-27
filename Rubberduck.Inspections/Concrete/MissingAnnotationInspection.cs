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
            Listener = new MissingAttributeAnnotationListener();
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;
        public IInspectionListener Listener { get; }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts.Select(context => new MissingAnnotationInspectionResult(this, context, context.MemberName));
        }

        public class MissingAttributeAnnotationListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly HashSet<string> _attributeNames;

            public MissingAttributeAnnotationListener()
            {
                _attributeNames = new HashSet<string>(typeof(AnnotationType).GetFields()
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

            #region scoping
            private IAnnotatedContext _currentScope;
            private string _currentScopeName;

            public override void EnterModuleBody(VBAParser.ModuleBodyContext context)
            {
                var firstMember = context.moduleBodyElement().FirstOrDefault()?.GetChild(0);
                _currentScope = firstMember as IAnnotatedContext;
                // name?
            }

            public override void ExitModuleAttributes(VBAParser.ModuleAttributesContext context)
            {
                if (_currentScope == null)
                {
                    // anything we pick up between here and the actual module body, belongs to the module
                    _currentScope = context;
                    _currentScopeName = CurrentModuleName.Name;
                }
                else
                {
                    // don't re-assign _currentScope here.
                    // we're at the end of the module and that attribute actually belongs to the last procedure.
                }
            }

            public override void EnterSubStmt(VBAParser.SubStmtContext context)
            {
                _currentScope = context;
            }

            public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
            {
                _currentScope = context;
            }

            public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
            {
                _currentScope = context;
            }

            public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
            {
                _currentScope = context;
            }

            public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
            {
                _currentScope = context;
            }
            #endregion

            public override void ExitAttributeStmt(VBAParser.AttributeStmtContext context)
            {
                if(_currentScope == null)
                {
                    // not scoped yet can't be a member attribute
                    return;
                }

                var name = context.attributeName().GetText();
                var value = context.attributeValue();
                if(!_currentScope.Annotations.Any(a => a.AnnotationType.HasFlag(AnnotationType.Attribute)
                                                       && _attributeNames.Select(n => $"{_currentScopeName}.{n}")
                                                                         .All(n => n != name)))
                {
                    // current scope is POSSIBLY missing an annotation for this attribute... todo: verify the value too
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }

                base.ExitAttributeStmt(context);
            }
        }
    }
}