using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.Command;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Navigation.CodeMetrics
{
    public class CodeMetricsViewModel : ViewModelBase
    {
        private readonly RubberduckParserState _state;

        public CodeMetricsViewModel(RubberduckParserState state, List<CommandBase> commands, ICodeMetricsAnalyst analyst)

        {
            _state = state;
            var reparseCommand = commands.OfType<ReparseCommand>().SingleOrDefault();
            RefreshCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(),
                reparseCommand == null ? (Action<object>)(o => { }) :
                o => reparseCommand.Execute(o),
                o => !IsBusy && reparseCommand != null && reparseCommand.CanExecute(o));
            
            _state.StateChanged += (_, change) =>
            {
                if (change.State == ParserState.Ready)
                {
                    ModuleMetrics = analyst.ModuleMetrics(_state);
                }
            };
        }

        public void FilterByName(object projects, string text)
        {
            throw new NotImplementedException();
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

        
        public CommandBase RefreshCommand { get; set; }


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
