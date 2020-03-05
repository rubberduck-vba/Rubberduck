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
    public abstract class ParseTreeInspectionBase<TContext> : InspectionBase, IParseTreeInspection
        where TContext : ParserRuleContext
    {
        protected ParseTreeInspectionBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        public IInspectionListener Listener => ContextListener;

        protected abstract IInspectionListener<TContext> ContextListener { get; }

        protected abstract string ResultDescription(QualifiedContext<TContext> context);

        protected virtual bool IsResultContext(QualifiedContext<TContext> context) => true;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return DoGetInspectionResults(ContextListener.Contexts());
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            return DoGetInspectionResults(ContextListener.Contexts(module));
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(IEnumerable<QualifiedContext<TContext>> contexts)
        {
            var objectionableContexts = contexts
                .Where(IsResultContext);

            return objectionableContexts
                .Select(InspectionResult)
                .ToList();
        }

        protected virtual IInspectionResult InspectionResult(QualifiedContext<TContext> context)
        {
            return new QualifiedContextInspectionResult(
                this,
                ResultDescription(context),
                context,
                DisabledQuickFixes(context));
        }

        protected virtual ICollection<string> DisabledQuickFixes(QualifiedContext<TContext> context) => new List<string>();
        public virtual CodeKind TargetKindOfCode => CodeKind.CodePaneCode;
    }


    public abstract class ParseTreeInspectionBase<TContext, TProperties> : InspectionBase, IParseTreeInspection
        where TContext : ParserRuleContext
    {
        protected ParseTreeInspectionBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        public IInspectionListener Listener => ContextListener;

        protected abstract IInspectionListener<TContext> ContextListener { get; }
        protected abstract string ResultDescription(QualifiedContext<TContext> context, TProperties properties);
        protected abstract (bool isResult, TProperties properties) IsResultContextWithAdditionalProperties(QualifiedContext<TContext> context);

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return DoGetInspectionResults(ContextListener.Contexts());
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            return DoGetInspectionResults(ContextListener.Contexts(module));
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(IEnumerable<QualifiedContext<TContext>> contexts)
        {
            var objectionableContexts = contexts
                .Select(ContextsWithResultProperties)
                .Where(result => result.HasValue)
                .Select(result => result.Value);

            return objectionableContexts
                .Select(tpl => InspectionResult(tpl.context, tpl.properties))
                .ToList();
        }

        private (QualifiedContext<TContext> context, TProperties properties)? ContextsWithResultProperties(QualifiedContext<TContext> context)
        {
            var (isResult, properties) = IsResultContextWithAdditionalProperties(context);
            return isResult
                ? (context, properties)
                : ((QualifiedContext<TContext> context, TProperties properties)?) null;
        }

        protected virtual IInspectionResult InspectionResult(QualifiedContext<TContext> context, TProperties properties)
        {
            return new QualifiedContextInspectionResult<TProperties>(
                this,
                ResultDescription(context, properties),
                context,
                properties,
                DisabledQuickFixes(context, properties));
        }

        protected virtual ICollection<string> DisabledQuickFixes(QualifiedContext<TContext> context, TProperties properties) => new List<string>();
        public virtual CodeKind TargetKindOfCode => CodeKind.CodePaneCode;
    }

    public class InspectionListenerBase<TContext> : VBAParserBaseListener, IInspectionListener<TContext>
        where TContext : ParserRuleContext
    {
        private readonly IDictionary<QualifiedModuleName, List<QualifiedContext<TContext>>> _contexts;

        public InspectionListenerBase()
        {
            _contexts = new Dictionary<QualifiedModuleName, List<QualifiedContext<TContext>>>();
        }

        public QualifiedModuleName CurrentModuleName { get; set; }
        
        public IReadOnlyList<QualifiedContext<TContext>> Contexts()
        {
            return _contexts.AllValues().ToList();
        }

        public IReadOnlyList<QualifiedContext<TContext>> Contexts(QualifiedModuleName module)
        {
            return _contexts.TryGetValue(module, out var contexts)
                ? contexts
                : new List<QualifiedContext<TContext>>();
        }

        public virtual void ClearContexts()
        {
            _contexts.Clear();
        }

        public virtual void ClearContexts(QualifiedModuleName module)
        {
            _contexts.Remove(module);
        }

        protected void SaveContext(TContext context)
        {
            var module = CurrentModuleName;
            var qualifiedContext = new QualifiedContext<TContext>(module, context);
            if (_contexts.TryGetValue(module, out var contexts))
            {
                contexts.Add(qualifiedContext);
            }
            else
            {
                _contexts.Add(module, new List<QualifiedContext<TContext>>{qualifiedContext});
            }
        }
    }
}