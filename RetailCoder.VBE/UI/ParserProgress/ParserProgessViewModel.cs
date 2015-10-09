using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.UI.ParserProgress
{
    public class ParserProgessViewModel : ViewModelBase
    {
        private readonly IRubberduckParser _parser;
        private readonly VBProject _project;

        public ParserProgessViewModel(IRubberduckParser parser, VBProject project)
        {
            _parser = parser;
            _project = project;
            var details = _project.VBComponents.Cast<VBComponent>().Select(component => new ComponentProgressViewModel(component)).ToList();
            _details = new ObservableCollection<ComponentProgressViewModel>(details);
        }

        private string _statusText;
        public string StatusText
        {
            get { return _statusText; }
            set
            {
                _statusText = value;
                OnPropertyChanged();
            }
        }

        private readonly ObservableCollection<ComponentProgressViewModel> _details;
        public IEnumerable<ComponentProgressViewModel> Details { get { return _details; } }

        public void Start()
        {
            _parser.ParseStarted += _parser_ParseStarted;
            _parser.ParseProgress += _parser_ParseProgress;
            _parser.ResolutionProgress += _parser_ResolutionProgress;
            
            _parser.Parse(_project);

            _parser.ResolutionProgress -= _parser_ResolutionProgress;
            _parser.ParseStarted -= _parser_ParseStarted;
            _parser.ParseProgress -= _parser_ParseProgress;
        }

        void _parser_ResolutionProgress(object sender, ResolutionProgressEventArgs e)
        {
            StatusText = "Resolving...";
        }

        void _parser_ParseProgress(object sender, ParseProgressEventArgs e)
        {
            StatusText = string.Format(RubberduckUI.ParseProgress, e.Component.Collection.Parent.Name + "." + e.Component.Name);
        }

        void _parser_ParseStarted(object sender, ParseStartedEventArgs e)
        {
            StatusText = RubberduckUI.ParseStarted;
        }
    }

    public class ComponentProgressViewModel : ViewModelBase
    {
        public ComponentProgressViewModel(VBComponent component)
        {
            ComponentName = component.Name;
        }

        public BitmapImage ComponentIcon { get; private set; } // todo: derive icon from component type
        public string ComponentName { get; private set; }

        private decimal _value;

        public decimal ResolutionProgressPercent
        {
            get { return _value; }
            set
            {
                _value = value;
                OnPropertyChanged();
            }
        }
    }
}
