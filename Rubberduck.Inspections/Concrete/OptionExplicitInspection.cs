using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class OptionExplicitInspection : InspectionBase, IParseTreeInspection
    {
        public OptionExplicitInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public IInspectionListener Listener { get; } = new MissingOptionExplicitListener();

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts.Select(context => new QualifiedContextInspectionResult(this,
                                                                      string.Format(InspectionsUI.OptionExplicitInspectionResultFormat, context.ModuleName.ComponentName),
                                                                      State,
                                                                      context));
        }

        public class MissingOptionExplicitListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitModuleDeclarations([NotNull] VBAParser.ModuleDeclarationsContext context)
            {
                var hasOptionExplicit = false;
                foreach (var element in context.moduleDeclarationsElement())
                {
                    if (element.moduleOption() is VBAParser.OptionExplicitStmtContext)
                    {
                        hasOptionExplicit = true;
                    }
                }

                if (!hasOptionExplicit)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, (ParserRuleContext)context.Parent));
                }
            }
        }
    }
}
