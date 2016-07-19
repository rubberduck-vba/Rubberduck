using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using NLog;
using Rubberduck.Inspections;
using Rubberduck.Settings;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class InspectionSettingsViewModel : ViewModelBase, ISettingsViewModel
    {
        public InspectionSettingsViewModel(Configuration config)
        {
            InspectionSettings = new ListCollectionView(
                    config.UserSettings.CodeInspectionSettings.CodeInspections.ToList());

            WhitelistedNameSettings = new ObservableCollection<WhitelistedNameSetting>(
                config.UserSettings.CodeInspectionSettings.WhitelistedNames.OrderBy(o => o.Name).Distinct());

            if (InspectionSettings.GroupDescriptions != null)
            {
                InspectionSettings.GroupDescriptions.Add(new PropertyGroupDescription("TypeLabel"));
            }
        }

        public void UpdateCollection(CodeInspectionSeverity severity)
        {
            // commit UI edit
            var item = (CodeInspectionSetting)InspectionSettings.CurrentEditItem;
            InspectionSettings.CommitEdit();

            // update the collection
            InspectionSettings.EditItem(item);
            item.Severity = severity;
            InspectionSettings.CommitEdit();
        }

        private ListCollectionView _inspectionSettings;
        public ListCollectionView InspectionSettings
        {
            get { return _inspectionSettings; }
            set
            {
                if (_inspectionSettings != value)
                {
                    _inspectionSettings = value;
                    OnPropertyChanged();
                }
            }
        }

        private ObservableCollection<WhitelistedNameSetting> _whitelistedNameSettings;
        public ObservableCollection<WhitelistedNameSetting> WhitelistedNameSettings
        {
            get { return _whitelistedNameSettings; }
            set
            {
                if (_whitelistedNameSettings != value)
                {
                    _whitelistedNameSettings = value;
                    OnPropertyChanged();
                }
            }
        }

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.CodeInspectionSettings.CodeInspections = new HashSet<CodeInspectionSetting>(InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>());
            config.UserSettings.CodeInspectionSettings.WhitelistedNames = WhitelistedNameSettings.Distinct().ToArray();
        }

        public void SetToDefaults(Configuration config)
        {
            InspectionSettings = new ListCollectionView(
                config.UserSettings.CodeInspectionSettings.CodeInspections.ToList());

            if (InspectionSettings.GroupDescriptions != null)
            {
                InspectionSettings.GroupDescriptions.Add(new PropertyGroupDescription("TypeLabel"));
            }

            WhitelistedNameSettings = new ObservableCollection<WhitelistedNameSetting>();
        }

        private CommandBase _addWhitelistedNameCommand;
        public CommandBase AddWhitelistedNameCommand
        {
            get
            {
                if (_addWhitelistedNameCommand != null)
                {
                    return _addWhitelistedNameCommand;
                }
                return _addWhitelistedNameCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ =>
                {
                    var placeholder = WhitelistedNameSettings.Count(m => m.Name.StartsWith("PLACEHOLDER")) + 1;
                    WhitelistedNameSettings.Add(
                        new WhitelistedNameSetting(string.Format("PLACEHOLDER{0}",
                            placeholder == 1 ? string.Empty : placeholder.ToString(CultureInfo.InvariantCulture))));
                });
            }
        }

        private CommandBase _deleteWhitelistedNameCommand;
        public CommandBase DeleteWhitelistedNameCommand
        {
            get
            {
                if (_deleteWhitelistedNameCommand != null)
                {
                    return _deleteWhitelistedNameCommand;
                }
                return _deleteWhitelistedNameCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), value =>
                {
                    WhitelistedNameSettings.Remove(value as WhitelistedNameSetting);
                });
            }
        }
    }
}
