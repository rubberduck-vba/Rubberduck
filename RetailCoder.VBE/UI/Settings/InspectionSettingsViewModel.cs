using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Data;
using NLog;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class InspectionSettingsViewModel : SettingsViewModelBase, ISettingsViewModel
    {
        public InspectionSettingsViewModel(Configuration config)
        {
            InspectionSettings = new ListCollectionView(
                    config.UserSettings.CodeInspectionSettings.CodeInspections.ToList());

            WhitelistedIdentifierSettings = new ObservableCollection<WhitelistedIdentifierSetting>(
                config.UserSettings.CodeInspectionSettings.WhitelistedIdentifiers.OrderBy(o => o.Identifier).Distinct());

            RunInspectionsOnSuccessfulParse = config.UserSettings.CodeInspectionSettings.RunInspectionsOnSuccessfulParse;

            if (InspectionSettings.GroupDescriptions != null)
            {
                InspectionSettings.GroupDescriptions.Add(new PropertyGroupDescription("TypeLabel"));
            }
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings());
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
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

        private bool _runInspectionsOnSuccessfulParse;

        public bool RunInspectionsOnSuccessfulParse
        {
            get { return _runInspectionsOnSuccessfulParse; }
            set
            {
                if (_runInspectionsOnSuccessfulParse != value)
                {
                    _runInspectionsOnSuccessfulParse = value;
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
            config.UserSettings.CodeInspectionSettings.RunInspectionsOnSuccessfulParse = _runInspectionsOnSuccessfulParse;
        }

        public void SetToDefaults(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.CodeInspectionSettings);
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

        private void TransferSettingsToView(CodeInspectionSettings toLoad)
        {
            InspectionSettings = new ListCollectionView(
                toLoad.CodeInspections.ToList());

            if (InspectionSettings.GroupDescriptions != null)
            {
                InspectionSettings.GroupDescriptions.Add(new PropertyGroupDescription("TypeLabel"));
            }

            WhitelistedIdentifierSettings = new ObservableCollection<WhitelistedIdentifierSetting>();
            RunInspectionsOnSuccessfulParse = true;
        }

        private void ImportSettings()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_LoadInspectionSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<CodeInspectionSettings> { FilePath = dialog.FileName };
                var loaded = service.Load(new CodeInspectionSettings());
                TransferSettingsToView(loaded);
            }
        }

        private void ExportSettings()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_SaveInspectionSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<CodeInspectionSettings> { FilePath = dialog.FileName };
                service.Save(new CodeInspectionSettings
                {
                    CodeInspections = new HashSet<CodeInspectionSetting>(InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>()),
                    WhitelistedIdentifiers = WhitelistedIdentifierSettings.Distinct().ToArray(),
                    RunInspectionsOnSuccessfulParse = _runInspectionsOnSuccessfulParse
                });
            }
        }
    }
}
