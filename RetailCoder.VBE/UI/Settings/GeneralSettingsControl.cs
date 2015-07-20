using System.Linq;
using System.Windows.Forms;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public partial class GeneralSettingsControl : UserControl
    {
        public GeneralSettingsControl()
        {
            InitializeComponent();
            TitleLabel.Text = RubberduckUI.SettingsCaption_GeneralSettings;
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
            var settings = new[]
            {
                new DisplayLanguageSetting("en-US"),
                new DisplayLanguageSetting("fr-CA"),
                new DisplayLanguageSetting("sv-SV"),
                new DisplayLanguageSetting("de-DE"),
            };

            // ReSharper disable once CoVariantArrayConversion
            LanguageList.Items.AddRange(settings.Where(item => item.Exists).ToArray());
            LanguageList.DisplayMember = "Name";
        }
    }
}
