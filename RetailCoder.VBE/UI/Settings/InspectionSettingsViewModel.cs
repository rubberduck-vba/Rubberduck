using System.Linq;
using System.Windows.Data;
using Rubberduck.Inspections;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class InspectionSetting
    {
        public string Name { get; set; }
        public CodeInspectionSeverity Severity { get; set; }
        public CodeInspectionType Type { get; set; }

        public InspectionSetting(CodeInspectionSetting setting)
        {
            Name = setting.Name;
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
                InspectionSettings.GroupDescriptions.Add(new PropertyGroupDescription("Type"));
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