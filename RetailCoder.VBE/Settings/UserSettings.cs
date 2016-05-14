using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    [XmlType(AnonymousType = true)]
    public class UserSettings
    {
        public GeneralSettings GeneralSettings { get; set; }
        public HotkeySettings HotkeySettings { get; set; }
        public ToDoListSettings ToDoListSettings { get; set; }
        public CodeInspectionSettings CodeInspectionSettings { get; set; }
        public UnitTestSettings UnitTestSettings { get; set; }
        public IndenterSettings IndenterSettings { get; set; }

        public UserSettings()
        {
            //default constructor required for serialization
        }

        public UserSettings(GeneralSettings generalSettings,
                            HotkeySettings hotkeySettings,
                            ToDoListSettings todoSettings, 
                            CodeInspectionSettings codeInspectionSettings,
                            UnitTestSettings unitTestSettings,
                            IndenterSettings indenterSettings)
        {
            GeneralSettings = generalSettings;
            HotkeySettings = hotkeySettings;
            ToDoListSettings = todoSettings;
            CodeInspectionSettings = codeInspectionSettings;
            UnitTestSettings = unitTestSettings;
            IndenterSettings = indenterSettings;
        }
    }
}
