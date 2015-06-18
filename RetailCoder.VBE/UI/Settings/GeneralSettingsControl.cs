using System.Windows.Forms;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public partial class GeneralSettingsControl : UserControl
    {
        private readonly IGeneralConfigService _configService;
        private readonly DisplayLanguageSetting _currentLanguage;

        public GeneralSettingsControl()
        {
            InitializeComponent();
            TitleLabel.Text = RubberduckUI.SettingsCaption_GeneralSettings;
            LanguageLabel.Text = RubberduckUI.Settings_LanguageLabel;

            LoadLanguageList();
        }

        public GeneralSettingsControl(DisplayLanguageSetting displayLanguage, IGeneralConfigService configService)
            : this()
        {
            _configService = configService;
            _currentLanguage = displayLanguage;
            LanguageList.SelectedItem = displayLanguage;
        }

        public DisplayLanguageSetting SelectedLanguage
        {
            get { return (DisplayLanguageSetting)LanguageList.SelectedItem; }
        }

        private void LoadLanguageList()
        {
            LanguageList.Items.Add(new DisplayLanguageSetting("en-US"));
            LanguageList.Items.Add(new DisplayLanguageSetting("fr-CA"));
            LanguageList.Items.Add(new DisplayLanguageSetting("sv-SV"));

            LanguageList.DisplayMember = "Name";
        }
    }
}
