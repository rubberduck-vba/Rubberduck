using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Results;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class ParseTreeInspectionBase : InspectionBase, IParseTreeInspection
    {
        protected ParseTreeInspectionBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        public abstract IInspectionListener Listener { get; }
        protected abstract string ResultDescription(QualifiedContext<ParserRuleContext> context);

        protected virtual bool IsResultContext(QualifiedContext<ParserRuleContext> context) => true;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return DoGetInspectionResults(Listener.Contexts());
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            return DoGetInspectionResults(Listener.Contexts(module));
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(IEnumerable<QualifiedContext<ParserRuleContext>> contexts)
        {
            var objectionableContexts = contexts
                .Where(IsResultContext);

            return objectionableContexts
                .Select(InspectionResult)
                .ToList();
        }

        protected virtual IInspectionResult InspectionResult(QualifiedContext<ParserRuleContext> context)
        {
            return new QualifiedContextInspectionResult(
                this,
                ResultDescription(context),
                context,
                DisabledQuickFixes(context));
        }

        protected virtual ICollection<string> DisabledQuickFixes(QualifiedContext<ParserRuleContext> context) => new List<string>();
        public virtual CodeKind TargetKindOfCode => CodeKind.CodePaneCode;
    }


    public abstract class ParseTreeInspectionBase<T> : InspectionBase, IParseTreeInspection
    {
        protected ParseTreeInspectionBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        public abstract IInspectionListener Listener { get; }
        protected abstract string ResultDescription(QualifiedContext<ParserRuleContext> context, T properties);
        protected abstract (bool isResult, T properties) IsResultContextWithAdditionalProperties(QualifiedContext<ParserRuleContext> context);

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return DoGetInspectionResults(Listener.Contexts());
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            return DoGetInspectionResults(Listener.Contexts(module));
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(IEnumerable<QualifiedContext<ParserRuleContext>> contexts)
        {
            var objectionableContexts = contexts
                .Select(ContextsWithResultProperties)
                .Where(result => result.HasValue)
                .Select(result => result.Value);

            return objectionableContexts
                .Select(tpl => InspectionResult(tpl.context, tpl.properties))
                .ToList();
        }

        private (QualifiedContext<ParserRuleContext> context, T properties)? ContextsWithResultProperties(QualifiedContext<ParserRuleContext> context)
        {
            var (isResult, properties) = IsResultContextWithAdditionalProperties(context);
            return isResult
                ? (context, properties)
                : ((QualifiedContext<ParserRuleContext> context, T properties)?) null;
        }

        protected virtual IInspectionResult InspectionResult(QualifiedContext<ParserRuleContext> context, T properties)
        {
            return new QualifiedContextInspectionResult<T>(
                this,
                ResultDescription(context, properties),
                context,
                properties,
                DisabledQuickFixes(context, properties));
        }

        protected virtual ICollection<string> DisabledQuickFixes(QualifiedContext<ParserRuleContext> context, T properties) => new List<string>();
        public virtual CodeKind TargetKindOfCode => CodeKind.CodePaneCode;
    }

    public class InspectionListenerBase : VBAParserBaseListener, IInspectionListener
    {
        private readonly IDictionary<QualifiedModuleName, List<QualifiedContext<ParserRuleContext>>> _contexts;

        public InspectionListenerBase()
        {
            _contexts = new Dictionary<QualifiedModuleName, List<QualifiedContext<ParserRuleContext>>>();
        }

        public QualifiedModuleName CurrentModuleName { get; set; }
        
        public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts()
        {
            return _contexts.AllValues().ToList();
        }

        public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts(QualifiedModuleName module)
        {
            return _contexts.TryGetValue(module, out var contexts)
                ? contexts
                : new List<QualifiedContext<ParserRuleContext>>();
        }

        public virtual void ClearContexts()
        {
            _contexts.Clear();
        }

        protected void SaveContext(ParserRuleContext context)
        {
            var module = CurrentModuleName;
            var qualifiedContext = new QualifiedContext<ParserRuleContext>(module, context);
            if (_contexts.TryGetValue(module, out var contexts))
            {
                contexts.Add(qualifiedContext);
            }
            else
            {
                _contexts.Add(module, new List<QualifiedContext<ParserRuleContext>>{qualifiedContext});
            }
        }
    }
}