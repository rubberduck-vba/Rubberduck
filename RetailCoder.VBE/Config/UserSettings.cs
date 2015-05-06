using System.Xml.Serialization;

namespace Rubberduck.Config
{
    [XmlType(AnonymousType = true)]
    public class UserSettings
    {
        public ToDoListSettings ToDoListSettings { get; set; }
        public CodeInspectionSettings CodeInspectionSettings { get; set; }

        public UserSettings()
        {
            //default constructor required for serialization
        }

        public UserSettings(ToDoListSettings todoSettings, CodeInspectionSettings codeInspectionSettings)
        {
            this.ToDoListSettings = todoSettings;
            this.CodeInspectionSettings = codeInspectionSettings;
        }
    }
}
