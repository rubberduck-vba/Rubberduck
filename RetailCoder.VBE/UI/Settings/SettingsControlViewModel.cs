using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Rubberduck.UI.Command;

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

        public event EventHandler OnOKButtonClicked;
        public event EventHandler OnCancelButtonClicked;

        private ICommand _okButtonCommand;
        public ICommand OKButtonCommand
        {
            get
            {
                if (_okButtonCommand != null)
                {
                    return _okButtonCommand;
                }
                return _okButtonCommand = new DelegateCommand(_ =>
                {
                    var handler = OnOKButtonClicked;
                    if (handler != null)
                    {
                        // todo update config
                        handler(this, EventArgs.Empty);
                    }
                });
            }
        }

        private ICommand _cancelButtonCommand;
        public ICommand CancelButtonCommand
        {
            get
            {
                if (_cancelButtonCommand != null)
                {
                    return _cancelButtonCommand;
                }
                return _cancelButtonCommand = new DelegateCommand(_ =>
                {
                    var handler = OnCancelButtonClicked;
                    if (handler != null)
                    {
                        handler(this, EventArgs.Empty);
                    }
                });
            }
        }

        private ICommand _refreshButtonCommand;
        public ICommand RefreshButtonCommand
        {
            get
            {
                if (_refreshButtonCommand != null)
                {
                    return _refreshButtonCommand;
                }
                return _refreshButtonCommand = new DelegateCommand(_ =>
                {
                    // todo impplement
                });
            }
        }
    }
}