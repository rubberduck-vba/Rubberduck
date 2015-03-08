using System.Runtime.InteropServices;
using System.Xml.Serialization;
using Rubberduck.Inspections;

namespace Rubberduck.Config
{
    [XmlTypeAttribute(AnonymousType = true)]
    public class CodeInspectionSettings
    {
        [XmlArrayItemAttribute("CodeInspection", IsNullable = false)]
        public CodeInspectionSetting[] CodeInspections { get; set; }

        public CodeInspectionSettings()
        {
            //default constructor requied for serialization
        }

        public CodeInspectionSettings(CodeInspectionSetting[] inspections)
        {
            this.CodeInspections = inspections;
        }
    }

    [XmlTypeAttribute(AnonymousType = true)]
    public class CodeInspectionSetting : IInspectionModel
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        public CodeInspectionSeverity Severity { get; set; }

        [XmlAttribute]
        public CodeInspectionType InspectionType { get; set; }

        public CodeInspectionSetting()
        {
            //default constructor required for serialization
        }

        public CodeInspectionSetting(string name, CodeInspectionType type, CodeInspectionSeverity severity)
        {
            this.Name = name;
            this.InspectionType = type;
            this.Severity = severity;
        }

        public CodeInspectionSetting(IInspectionModel inspection)
            : this(inspection.Name, inspection.InspectionType, inspection.Severity)
        { }
    }
}
