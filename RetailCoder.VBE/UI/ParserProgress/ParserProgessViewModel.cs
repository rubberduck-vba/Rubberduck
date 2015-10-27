using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Properties;

namespace Rubberduck.UI.ParserProgress
{
    public class ParserProgessViewModel : ViewModelBase
    {
        private readonly IRubberduckParser _parser;
        private readonly VBProject _project;

        public ParserProgessViewModel(IRubberduckParser parser, VBProject project)
        {
            _parser = parser;
            _parser.ParseStarted += _parser_ParseStarted;
            _parser.ParseProgress += _parser_ParseProgress;
            _parser.ResolutionProgress += _parser_ResolutionProgress;
            _parser.ResolutionCompleted += _parser_Completed;

            _project = project;
            var details = _project.VBComponents.Cast<VBComponent>().Select(component => new ComponentProgressViewModel(component)).ToList();
            _details = new ObservableCollection<ComponentProgressViewModel>(details);
        }

        public IRubberduckParser Parser { get { return _parser; } }

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
        public ObservableCollection<ComponentProgressViewModel> Details { get { return _details; } }

        public void Start()
        {            
            _parser.Parse(_project, this);
        }

        public event EventHandler<EventArgs> Completed;
        void _parser_Completed(object sender, EventArgs e)
        {
            var handler = Completed;
            if (handler != null)
            {
                handler.Invoke(this, e);
            }
        }

        void _parser_ResolutionProgress(object sender, ResolutionProgressEventArgs e)
        {
            StatusText = RubberduckUI.ResolutionProgress;
            var row = _details.SingleOrDefault(vm => vm.ComponentName == e.Component.Name);
            if (row == null)
            {
                return;
            }
            row.ResolutionProgressPercent = e.PercentProgress;
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
        private readonly VBComponent _component;

        public ComponentProgressViewModel(VBComponent component)
        {
            _component = component;
            ComponentName = component.Name;
        }

        public BitmapImage ComponentIcon
        {
            get { return DeclarationIconCache.ComponentIcon(_component.Type); }
        }

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
