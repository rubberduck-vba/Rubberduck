using System.Xml.Serialization;
using Rubberduck.SmartIndenter;

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
        public WindowSettings WindowSettings { get; set; }

        public UserSettings(GeneralSettings generalSettings,
                            HotkeySettings hotkeySettings,
                            ToDoListSettings todoSettings, 
                            CodeInspectionSettings codeInspectionSettings,
                            UnitTestSettings unitTestSettings,
                            IndenterSettings indenterSettings,
                            WindowSettings windowSettings)
        {
            GeneralSettings = generalSettings;
            HotkeySettings = hotkeySettings;
            ToDoListSettings = todoSettings;
            CodeInspectionSettings = codeInspectionSettings;
            UnitTestSettings = unitTestSettings;
            IndenterSettings = indenterSettings;
            WindowSettings = windowSettings;
        }
    }
}
