using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    internal class NestingLevelMetric : CodeMetric
    {
        public NestingLevelMetric() : base("Nesting Level", AggregationLevel.Member) { }

        public override ICodeMetricsParseTreeListener TreeListener => new NestingLevelListener(this);
    }

    internal class NestingLevelMetricResult : CodeMetricResultBase
    {
        private readonly int value;
        public NestingLevelMetricResult(Declaration declaration, CodeMetric metric, int value) : base (declaration, metric)
        {
            this.value = value;
        }

        public override string Value => value.ToString();
    }

    internal class NestingLevelListener : CodeMetricListenerBase
    {
        private List<ICodeMetricResult> _results = new List<ICodeMetricResult>();
        private int _currentNestingLevel;
        private int _currentMaxNesting;

        public NestingLevelListener(CodeMetric owner) : base(owner)
        {
        }

        public override void Reset()
        {
            base.Reset();
            _results = new List<ICodeMetricResult>();
            _currentMaxNesting = _currentNestingLevel = 0;
        }

        public override IEnumerable<ICodeMetricResult> Results() => _results;

        public override void EnterBlock([NotNull] VBAParser.BlockContext context)
        {
            _currentNestingLevel++;
            if (_currentNestingLevel > _currentMaxNesting)
            {
                _currentMaxNesting = _currentNestingLevel;
            }
        }

        public override void ExitBlock([NotNull] VBAParser.BlockContext context) => _currentNestingLevel--;
        public override void ExitPropertySetStmt([NotNull] VBAParser.PropertySetStmtContext context) 
            => ExitMeasurableMember(_finder.UserDeclarations(DeclarationType.PropertySet).Where(d => d.Context == context).First());
        public override void ExitPropertyLetStmt([NotNull] VBAParser.PropertyLetStmtContext context) 
            => ExitMeasurableMember(_finder.UserDeclarations(DeclarationType.PropertyLet).Where(d => d.Context == context).First());
        public override void ExitPropertyGetStmt([NotNull] VBAParser.PropertyGetStmtContext context) 
            => ExitMeasurableMember(_finder.UserDeclarations(DeclarationType.PropertyGet).Where(d => d.Context == context).First());
        public override void ExitFunctionStmt([NotNull] VBAParser.FunctionStmtContext context) 
            => ExitMeasurableMember(_finder.UserDeclarations(DeclarationType.Function).Where(d => d.Context == context).First());
        public override void ExitSubStmt([NotNull] VBAParser.SubStmtContext context) 
            => ExitMeasurableMember(_finder.UserDeclarations(DeclarationType.Procedure).Where(d => d.Context == context).First());
        private void ExitMeasurableMember(Declaration declaration)
        {
            Debug.Assert(_currentNestingLevel == 0, "Unexpected nesting level when exiting measurable member");
            _results.Add(new CyclomaticComplexityMetricResult(declaration, ownerReference, _currentMaxNesting));
            _currentMaxNesting = _currentNestingLevel = 0;
        }
    }
}
