using System.Xml.Serialization;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Settings
{
    [XmlType(AnonymousType = true)]
    public class UserSettings
    {
        public DisplayLanguageSetting LanguageSetting { get; set; }
        public ToDoListSettings ToDoListSettings { get; set; }
        public CodeInspectionSettings CodeInspectionSettings { get; set; }
        public IndenterSettings IndenterSettings { get; set; }

        public UserSettings()
        {
            //default constructor required for serialization
        }

        public UserSettings(DisplayLanguageSetting languageSetting,
                            ToDoListSettings todoSettings, 
                            CodeInspectionSettings codeInspectionSettings,
                            IndenterSettings indenterSettings)
        {
            LanguageSetting = languageSetting;
            ToDoListSettings = todoSettings;
            CodeInspectionSettings = codeInspectionSettings;
            IndenterSettings = indenterSettings;
        }
    }
}
