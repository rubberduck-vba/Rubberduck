using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Results;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class DefTypeStatementInspection : ParseTreeInspectionBase
    {
        public DefTypeStatementInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new DefTypeStatementInspectionListener();
        }
        
        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var results = Listener.Contexts.Where(context => !IsIgnoringInspectionResultFor(context.ModuleName, context.Context.Start.Line))
                .Select(context => new QualifiedContextInspectionResult(this,
                                                                        string.Format(InspectionsUI.DefTypeStatementInspectionResultFormat,
                                                                                      GetTypeOfDefType(context.Context.start.Text),
                                                                                      context.Context.start.Text),
                                                                        context));
            return results;
        }

        public class DefTypeStatementInspectionListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitDefType([NotNull] VBAParser.DefTypeContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }

        private string GetTypeOfDefType(string defType)
        {
            _defTypes.TryGetValue(defType, out var value);
            return value;
        }

        private readonly Dictionary<string, string> _defTypes = new Dictionary<string, string>
        {
            { "DefBool", "Boolean" },
            { "DefByte", "Byte" },
            { "DefInt", "Integer" },
            { "DefLng", "Long" },
            { "DefCur", "Currency" },
            { "DefSng", "Single" },
            { "DefDbl", "Double" },
            { "DefDate", "Date" },
            { "DefStr", "String" },
            { "DefObj", "Object" },
            { "DefVar", "Variant" }
        };
    }
}