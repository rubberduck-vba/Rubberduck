using System.Windows.Forms;
using Rubberduck.Config;

namespace Rubberduck.UI.Settings
{
    public partial class GeneralSettingsControl : UserControl
    {
        public GeneralSettingsControl()
        {
            InitializeComponent();
            TitleLabel.Text = RubberduckUI.SettingsCaption_GeneralSettings;
            InstructionsLabel.Text = RubberduckUI.SettingsInstructions_GeneralSettings;
            LanguageLabel.Text = RubberduckUI.Settings_LanguageLabel;

            LoadLanguageList();
        }

        public GeneralSettingsControl(DisplayLanguageSetting displayLanguage)
            : this()
        {
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

            LanguageList.DisplayMember = "Name";
        }
    }
}
