using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

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
                ModuleMetrics = new ObservableCollection<ModuleMetricsResult>(_analyst.ModuleMetrics(_state));
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

        private ObservableCollection<ModuleMetricsResult> _moduleMetrics;
        public ObservableCollection<ModuleMetricsResult> ModuleMetrics {
            get => _moduleMetrics;
            private set
            {
                _moduleMetrics = value;
                SelectedMetric = ModuleMetrics.Any(i => SelectedMetric.ModuleName == i.ModuleName)
                    ? ModuleMetrics.First(i => SelectedMetric.ModuleName == i.ModuleName)
                    : ModuleMetrics.FirstOrDefault();
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
