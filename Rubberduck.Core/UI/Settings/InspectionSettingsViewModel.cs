using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Data;
using NLog;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using Rubberduck.Resources.Inspections;
using System.Globalization;
using System;
using Rubberduck.Resources.Settings;

namespace Rubberduck.UI.Settings
{
    public class InspectionSettingsViewModel : SettingsViewModelBase, ISettingsViewModel
    {
        public InspectionSettingsViewModel(Configuration config)
        {
            InspectionSettings = new ListCollectionView(
                        config.UserSettings.CodeInspectionSettings.CodeInspections
                                        .OrderBy(inspection => inspection.TypeLabel)
                                        .ThenBy(inspection => inspection.Description)
                                        .ToList());

            WhitelistedIdentifierSettings = new ObservableCollection<WhitelistedIdentifierSetting>(
                config.UserSettings.CodeInspectionSettings.WhitelistedIdentifiers.OrderBy(o => o.Identifier).Distinct());

            RunInspectionsOnSuccessfulParse = config.UserSettings.CodeInspectionSettings.RunInspectionsOnSuccessfulParse;

            InspectionSettings.GroupDescriptions?.Add(new PropertyGroupDescription("TypeLabel"));
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings());
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());

            SeverityFilters = new ObservableCollection<string>(
                new[] { InspectionsUI.ResourceManager.GetString("CodeInspectionSeverity_All", CultureInfo.CurrentUICulture) }
                    .Concat(Enum.GetNames(typeof(CodeInspectionSeverity)).Select(s => InspectionsUI.ResourceManager.GetString("CodeInspectionSeverity_" + s, CultureInfo.CurrentUICulture))));
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

        private string _inspectionSettingsDescriptionFilter = string.Empty;
        public string InspectionSettingsDescriptionFilter
        {
            get => _inspectionSettingsDescriptionFilter;
            set
            {
                if (_inspectionSettingsDescriptionFilter != value)
                {
                    _inspectionSettingsDescriptionFilter = value;
                    InspectionSettings.Filter = item => FilterResults(item);
                }
            }
        }

        public ObservableCollection<string> SeverityFilters { get; }

        static private readonly string _allResultsFilter = InspectionsUI.ResourceManager.GetString("CodeInspectionSeverity_All", CultureInfo.CurrentUICulture);
        private string _selectedSeverityFilter = _allResultsFilter;
        public string SelectedSeverityFilter
        {
            get => _selectedSeverityFilter;
            set
            {
                if (!_selectedSeverityFilter.Equals(value))
                {
                    _selectedSeverityFilter = value.Replace(" ", string.Empty);
                    OnPropertyChanged();
                    InspectionSettings.Filter = item => FilterResults(item);
                }
            }
        }        

        private bool FilterResults(object setting)
        {
            OnPropertyChanged(nameof(InspectionSettings));
            var cis = setting as CodeInspectionSetting;

            return cis.Description.ToUpper().Contains(_inspectionSettingsDescriptionFilter.ToUpper())
                && (_selectedSeverityFilter.Equals(_allResultsFilter) || cis.Severity.ToString().Equals(_selectedSeverityFilter));
        }

        private ListCollectionView _inspectionSettings;
        public ListCollectionView InspectionSettings
        {
            get
            {
                return _inspectionSettings;
            }

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
            get => _runInspectionsOnSuccessfulParse;
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
            get => _whitelistedNameSettings;
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
            InspectionSettings = new ListCollectionView(toLoad.CodeInspections.ToList());
            
            InspectionSettings.GroupDescriptions.Add(new PropertyGroupDescription("TypeLabel"));
            
            WhitelistedIdentifierSettings = new ObservableCollection<WhitelistedIdentifierSetting>();
            RunInspectionsOnSuccessfulParse = true;
        }

        private void ImportSettings()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = SettingsUI.DialogMask_XmlFilesOnly,
                Title = SettingsUI.DialogCaption_LoadInspectionSettings
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
                Filter = SettingsUI.DialogMask_XmlFilesOnly,
                Title = SettingsUI.DialogCaption_SaveInspectionSettings
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
