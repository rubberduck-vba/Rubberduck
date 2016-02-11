using System.Linq;
using System.Windows.Data;
using Rubberduck.Inspections;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class InspectionSetting
    {
        public string Name { get; private set; }
        public string Description { get; set; }
        public CodeInspectionSeverity Severity { get; set; }
        public CodeInspectionType Type { get; set; }
        public string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }

        public string TypeLabel
        {
            get { return RubberduckUI.ResourceManager.GetString("CodeInspectionSettings_" + Type); }
        }

        public InspectionSetting(CodeInspectionSetting setting)
        {
            Name = setting.Name;
            Description = setting.Description;
            Severity = setting.Severity;
            Type = setting.InspectionType;
        }
    }

    public class InspectionSettingsViewModel : ViewModelBase
    {
        private readonly IGeneralConfigService _configService;
        private readonly Configuration _config;

        public InspectionSettingsViewModel(IGeneralConfigService configService)
        {
            _configService = configService;
            _config = configService.LoadConfiguration();

            InspectionSettings = new ListCollectionView(
                    _config.UserSettings.CodeInspectionSettings.CodeInspections.Select(i => new InspectionSetting(i))
                        .ToList());

            if (InspectionSettings.GroupDescriptions != null)
            {
                InspectionSettings.GroupDescriptions.Add(new PropertyGroupDescription("TypeLabel"));
            }
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
    }
}