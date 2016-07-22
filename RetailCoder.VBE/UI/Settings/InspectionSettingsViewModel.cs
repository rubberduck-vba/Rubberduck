using System.Collections.Generic;
using System.Collections.ObjectModel;
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

            WhitelistedIdentifierSettings = new ObservableCollection<WhitelistedIdentifierSetting>(
                config.UserSettings.CodeInspectionSettings.WhitelistedIdentifiers.OrderBy(o => o.Identifier).Distinct());

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

        private ObservableCollection<WhitelistedIdentifierSetting> _whitelistedNameSettings;
        public ObservableCollection<WhitelistedIdentifierSetting> WhitelistedIdentifierSettings
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
            config.UserSettings.CodeInspectionSettings.WhitelistedIdentifiers = WhitelistedIdentifierSettings.Distinct().ToArray();
        }

        public void SetToDefaults(Configuration config)
        {
            InspectionSettings = new ListCollectionView(
                config.UserSettings.CodeInspectionSettings.CodeInspections.ToList());

            if (InspectionSettings.GroupDescriptions != null)
            {
                InspectionSettings.GroupDescriptions.Add(new PropertyGroupDescription("TypeLabel"));
            }

            WhitelistedIdentifierSettings = new ObservableCollection<WhitelistedIdentifierSetting>();
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
                    WhitelistedIdentifierSettings.Add(new WhitelistedIdentifierSetting());
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
                    WhitelistedIdentifierSettings.Remove(value as WhitelistedIdentifierSetting);
                });
            }
        }
    }
}
