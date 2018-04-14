using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    internal class CyclomaticComplexityMemberMetric : CodeMetric
    {
        public CyclomaticComplexityMemberMetric() : base("Cyclomatic Complexity", AggregationLevel.Member)
        {
        }

        public override ICodeMetricsParseTreeListener TreeListener
        {
            get => new CyclomaticComplexityListener(this);
        }
    }

    internal class CyclomaticComplexityMetricResult : ICodeMetricResult
    {
        private readonly CodeMetric metricReference;
        private readonly Declaration declaration;
        private readonly int value;
        public CyclomaticComplexityMetricResult(Declaration declaration, CodeMetric metricReference, int value)
        {
            this.declaration = declaration;
            this.metricReference = metricReference;
            this.value = value;
        }

        public Declaration Declaration => declaration;

        public CodeMetric Metric => metricReference;

        public string Value => value.ToString();
    }

    internal class CyclomaticComplexityListener : VBAParserBaseListener, ICodeMetricsParseTreeListener
    {
        private readonly CodeMetric ownerReference;
        private List<ICodeMetricResult> _results = new List<ICodeMetricResult>();
        private DeclarationFinder _finder;
        private QualifiedModuleName _qmn;

        private int workingValue;

        public CyclomaticComplexityListener(CodeMetric owner)
        {
            ownerReference = owner;
        }

        public void InjectContext((DeclarationFinder, QualifiedModuleName) context) => (_finder, _qmn) = context;

        public void Reset()
        {
            _results = new List<ICodeMetricResult>();
            _finder = null;
            _qmn = default;
        }

        public IEnumerable<ICodeMetricResult> Results() => _results;

        public override void EnterIfStmt([NotNull] VBAParser.IfStmtContext context) => workingValue++;
        public override void EnterElseIfBlock([NotNull] VBAParser.ElseIfBlockContext context) => workingValue++;
        public override void EnterForEachStmt([NotNull] VBAParser.ForEachStmtContext context) => workingValue++;
        public override void EnterForNextStmt([NotNull] VBAParser.ForNextStmtContext context) => workingValue++;
        public override void EnterCaseClause([NotNull] VBAParser.CaseClauseContext context) => workingValue++;

        public override void ExitPropertySetStmt([NotNull] VBAParser.PropertySetStmtContext context)
        {
            ExitMeasurableMember(_finder.UserDeclarations(DeclarationType.PropertySet).Where(d => d.Context == context).First());
        }
        public override void ExitPropertyLetStmt([NotNull] VBAParser.PropertyLetStmtContext context)
        {
            ExitMeasurableMember(_finder.UserDeclarations(DeclarationType.PropertyLet).Where(d => d.Context == context).First());
        }
        public override void ExitPropertyGetStmt([NotNull] VBAParser.PropertyGetStmtContext context)
        {
            ExitMeasurableMember(_finder.UserDeclarations(DeclarationType.PropertyGet).Where(d => d.Context == context).First());
        }
        public override void ExitFunctionStmt([NotNull] VBAParser.FunctionStmtContext context)
        {
            ExitMeasurableMember(_finder.UserDeclarations(DeclarationType.Function).Where(d => d.Context == context).First());
        }

        public override void ExitSubStmt([NotNull] VBAParser.SubStmtContext context)
        {
            ExitMeasurableMember(_finder.UserDeclarations(DeclarationType.Procedure).Where(d => d.Context == context).First());
        }

        private void ExitMeasurableMember(Declaration declaration)
        {
            // handle enter of this member
            workingValue++;
            _results.Add(new CyclomaticComplexityMetricResult(declaration, ownerReference, workingValue));
            // reset working value
            workingValue = 0;
        }
    }

}
