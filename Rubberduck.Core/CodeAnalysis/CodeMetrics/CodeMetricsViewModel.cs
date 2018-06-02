using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using System;
using System.Linq;
using System.Collections.Generic;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    public class CodeMetricsViewModel : ViewModelBase, IDisposable
    {
        private readonly RubberduckParserState _state;
        private readonly ICodeMetricsAnalyst _analyst;

        public CodeMetricsViewModel(RubberduckParserState state, ICodeMetricsAnalyst analyst)
        {
            _state = state;
            _analyst = analyst;
            _state.StateChanged += OnStateChanged;
        }

        private void OnStateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State == ParserState.Ready)
            {
                IsBusy = true;

                var metricResults = _analyst.GetMetrics(_state);

                metricResultContainer = metricResults
                    .GroupBy(r => r.Metric.Level)
                    .ToDictionary(g => g.Key,
                       levelGrouping => levelGrouping.GroupBy(r => r.Declaration)
                         .ToDictionary(g => g.Key,
                            declarationGrouping => declarationGrouping.ToDictionary(r => r.Metric)
                         )
                    );

                metricsByLevel = metricResults.GroupBy(r => r.Metric.Level).ToDictionary(g => g.Key, g => g.Select(r => r.Metric).ToList());
                declarationsByLevel = metricResults.GroupBy(r => r.Metric.Level).ToDictionary(g => g.Key, g => g.Select(r => r.Declaration).ToList());
                declarationsByParent = metricResults.Select(r => r.Declaration).GroupBy(decl => decl.ParentDeclaration).ToDictionary(g => g.Key, g => g.ToList());
                resultsByDeclaration = metricResults.GroupBy(r => r.Declaration).ToDictionary(g => g.Key, g => g.ToList());
                

                IsBusy = false;
            }
        }

        public void Dispose()
        {
            _state.StateChanged -= OnStateChanged;
        }

        // TBD: use these dictionaries to populate the GridView
        private Dictionary<AggregationLevel, List<CodeMetric>> metricsByLevel;
        private Dictionary<AggregationLevel, List<Declaration>> declarationsByLevel;
        private Dictionary<Declaration, List<Declaration>> declarationsByParent;
        private Dictionary<Declaration, List<ICodeMetricResult>> resultsByDeclaration; 

        private Dictionary<AggregationLevel, Dictionary<Declaration, Dictionary<CodeMetric, ICodeMetricResult>>>
            metricResultContainer;
        public Dictionary<AggregationLevel, Dictionary<Declaration, Dictionary<CodeMetric, ICodeMetricResult>>>
            MetricResults
        { get => metricResultContainer; }

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                _isBusy = value;
                OnPropertyChanged();
                // If the window is "busy" then hide the Refresh message
                OnPropertyChanged("EmptyUIRefreshMessageVisibility");
            }
        }
    }
}
