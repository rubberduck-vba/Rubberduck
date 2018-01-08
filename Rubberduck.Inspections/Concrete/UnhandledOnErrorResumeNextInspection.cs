using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public class UnhandledOnErrorResumeNextInspection : ParseTreeInspectionBase
    {
        private readonly Dictionary<QualifiedContext<ParserRuleContext>, string> _errorHandlerLabelsMap =
            new Dictionary<QualifiedContext<ParserRuleContext>, string>();
        private readonly Dictionary<QualifiedContext<ParserRuleContext>, VBAParser.ModuleBodyElementContext> _bodyElementContextsMap =
            new Dictionary<QualifiedContext<ParserRuleContext>, VBAParser.ModuleBodyElementContext>();

        public UnhandledOnErrorResumeNextInspection(RubberduckParserState state,
            CodeInspectionSeverity defaultSeverity = CodeInspectionSeverity.Warning) : base(state, defaultSeverity)
        {
            Listener = new OnErrorStatementListener(_errorHandlerLabelsMap, _bodyElementContextsMap);
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result =>
                {
                    dynamic properties = new PropertyBag();
                    properties.Label = _errorHandlerLabelsMap[result];
                    properties.BodyElement = _bodyElementContextsMap[result];

                    return new QualifiedContextInspectionResult(this, InspectionsUI.UnhandledOnErrorResumeNextInspectionResultFormat, result, properties);
                });
        }
    }

    public class OnErrorStatementListener : VBAParserBaseListener, IInspectionListener
    {
        private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
        private readonly List<QualifiedContext<ParserRuleContext>> _unhandledContexts = new List<QualifiedContext<ParserRuleContext>>();
        private readonly List<string> _errorHandlerLabels = new List<string>();
        private readonly Dictionary<QualifiedContext<ParserRuleContext>, string> _errorHandlerLabelsMap;
        private readonly Dictionary<QualifiedContext<ParserRuleContext>, VBAParser.ModuleBodyElementContext> _bodyElementContextsMap;

        private const string LabelPrefix = "ErrorHandler";

        public OnErrorStatementListener(Dictionary<QualifiedContext<ParserRuleContext>, string> errorHandlerLabelsMap,
            Dictionary<QualifiedContext<ParserRuleContext>, VBAParser.ModuleBodyElementContext> bodyElementContextsMap)
        {
            _errorHandlerLabelsMap = errorHandlerLabelsMap;
            _bodyElementContextsMap = bodyElementContextsMap;
        }

        public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

        public void ClearContexts()
        {
            _contexts.Clear();
        }

        public QualifiedModuleName CurrentModuleName { get; set; }

        public override void ExitModuleBodyElement(VBAParser.ModuleBodyElementContext context)
        {
            if (_unhandledContexts.Any())
            {
                var labelIndex = -1;

                foreach (var errorContext in _unhandledContexts)
                {
                    _bodyElementContextsMap.Add(errorContext, context);

                    labelIndex++;
                    var labelSuffix = labelIndex == 0 ? "" : labelIndex.ToString();

                    while (_errorHandlerLabels.Contains($"{LabelPrefix.ToLower()}{labelSuffix}"))
                    {
                        labelIndex++;
                        labelSuffix = labelIndex == 0 ? "" : labelIndex.ToString();
                    }

                    _errorHandlerLabelsMap.Add(errorContext, $"{LabelPrefix}{labelSuffix}");
                }

                _contexts.AddRange(_unhandledContexts);

                _unhandledContexts.Clear();
                _errorHandlerLabels.Clear();
            }
        }

        public override void ExitOnErrorStmt(VBAParser.OnErrorStmtContext context)
        {
            if (context.RESUME() != null)
            {
                _unhandledContexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
            else if (context.GOTO() != null)
            {
                _unhandledContexts.Clear();
            }
        }

        public override void ExitIdentifierStatementLabel(VBAParser.IdentifierStatementLabelContext context)
        {
            var labelText = context.unrestrictedIdentifier().identifier().untypedIdentifier().identifierValue().IDENTIFIER().GetText();
            if (labelText.ToLower().StartsWith(LabelPrefix.ToLower()))
            {
                _errorHandlerLabels.Add(labelText.ToLower());
            }
        }
    }
}
