using System.Collections.ObjectModel;
using System.Linq;

namespace Rubberduck.UI.Settings
{
    public class SettingsView
    {
        public string Label { get; set; }
        public ISettingsView Control { get; set; }
        public SettingsViews View { get; set; }

        public SettingsView(string label, ISettingsView control, SettingsViews view)
        {
            Label = label;
            Control = control;
            View = view;
        }
    }

    public class SettingsControlViewModel : ViewModelBase
    {
        public SettingsControlViewModel(SettingsViews view = Settings.SettingsViews.GeneralSettings)
        {
            SettingsViews = new ObservableCollection<SettingsView>
            {
                new SettingsView(RubberduckUI.SettingsCaption_GeneralSettings, new GeneralSettings(), Settings.SettingsViews.GeneralSettings),
                new SettingsView(RubberduckUI.SettingsCaption_ToDoSettings, new GeneralSettings(), Settings.SettingsViews.TodoSettings),
                new SettingsView(RubberduckUI.SettingsCaption_CodeInspections, new GeneralSettings(), Settings.SettingsViews.InspectionSettings),
                new SettingsView(RubberduckUI.SettingsCaption_UnitTestSettings, new GeneralSettings(), Settings.SettingsViews.UnitTestSettings)
            };

            SelectedSettingsView = SettingsViews.First(v => v.View == view);
        }

        private ObservableCollection<SettingsView> _settingsViews;
        public ObservableCollection<SettingsView> SettingsViews
        {
            get
            {
                return _settingsViews;
            }
            set
            {
                if (_settingsViews != value)
                {
                    _settingsViews = value;
                    OnPropertyChanged();
                }
            }
        }

        private SettingsView _seletedSettingsView;
        public SettingsView SelectedSettingsView
        {
            get { return _seletedSettingsView; }
            set
            {
                if (_seletedSettingsView != value)
                {
                    _seletedSettingsView = value;
                    OnPropertyChanged();
                }
            }
        }
    }
}