using System.Linq;
using System.Windows.Data;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class InspectionSettingsViewModel : ViewModelBase, ISettingsViewModel
    {
        private readonly Configuration _config;

        public InspectionSettingsViewModel(Configuration config)
        {
            _config = config;

            InspectionSettings = new ListCollectionView(
                    _config.UserSettings.CodeInspectionSettings.CodeInspections.ToList());

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

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.CodeInspectionSettings.CodeInspections =
                InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>().ToArray();
        }
    }
}