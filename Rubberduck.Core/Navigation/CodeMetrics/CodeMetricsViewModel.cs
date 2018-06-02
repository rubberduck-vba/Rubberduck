using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using System;
using System.Collections.Generic;

namespace Rubberduck.Navigation.CodeMetrics
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
            if (e.State != ParserState.Ready && e.State != ParserState.Error && e.State != ParserState.ResolverError && e.State != ParserState.UnexpectedError)
            {
                IsBusy = true;
            }

            if (e.State == ParserState.Ready)
            {
                ModuleMetrics = _analyst.ModuleMetrics(_state);
                IsBusy = false;
            }

            if  (e.State == ParserState.Error || e.State == ParserState.ResolverError || e.State == ParserState.UnexpectedError)
            {
                IsBusy = false;
            }
        }

        public void Dispose()
        {
            _state.StateChanged -= OnStateChanged;
        }

        private ModuleMetricsResult _selectedMetric;
        public ModuleMetricsResult SelectedMetric
        {
            get => _selectedMetric;
            set
            {
                _selectedMetric = value;
                OnPropertyChanged();
            }
        }

        private IEnumerable<ModuleMetricsResult> _moduleMetrics;
        public IEnumerable<ModuleMetricsResult> ModuleMetrics {
            get => _moduleMetrics;
            private set
            {
                _moduleMetrics = value;
                OnPropertyChanged();
            }
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                _isBusy = value;
                EmptyUIRefreshMessageVisibility = false;
                OnPropertyChanged();
            }
        }


        private bool _emptyUIRefreshMessageVisibility = true;
        public bool EmptyUIRefreshMessageVisibility
        {
            get => _emptyUIRefreshMessageVisibility;
            set
            {
                if (_emptyUIRefreshMessageVisibility != value)
                {
                    _emptyUIRefreshMessageVisibility = value;
                    OnPropertyChanged();
                }
            }
        }
    }
}
