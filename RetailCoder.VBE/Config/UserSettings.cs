using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Runtime.InteropServices;

namespace Rubberduck.Config
{
    [XmlTypeAttribute(AnonymousType = true)]
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
