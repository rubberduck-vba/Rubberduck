using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.Command;
using System;
using System.Collections.Generic;
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
            if (e.State == ParserState.Ready)
            {
                ModuleMetrics = _analyst.ModuleMetrics(_state);
            }
        }

        public void FilterByName(object projects, string text)
        {
            throw new NotImplementedException();
        }

        public void Dispose()
        {
            _state.StateChanged -= OnStateChanged;
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

        
        //public CommandBase RefreshCommand { get; set; }


        private bool _canSearch;
        public bool CanSearch
        {
            get => _canSearch;
            set
            {
                _canSearch = value;
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
                OnPropertyChanged();
                // If the window is "busy" then hide the Refresh message
                OnPropertyChanged("EmptyUIRefreshMessageVisibility");
            }
        }
    }
}
