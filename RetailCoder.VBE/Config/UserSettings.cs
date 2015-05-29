using System.Xml.Serialization;

namespace Rubberduck.Config
{
    [XmlType(AnonymousType = true)]
    public class UserSettings
    {
        public DisplayLanguageSetting LanguageSetting { get; set; }
        public ToDoListSettings ToDoListSettings { get; set; }
        public CodeInspectionSettings CodeInspectionSettings { get; set; }

        public UserSettings()
        {
            //default constructor required for serialization
        }

        public UserSettings(DisplayLanguageSetting languageSetting,
                            ToDoListSettings todoSettings, 
                            CodeInspectionSettings codeInspectionSettings)
        {
            LanguageSetting = languageSetting;
            ToDoListSettings = todoSettings;
            CodeInspectionSettings = codeInspectionSettings;
        }
    }
}
