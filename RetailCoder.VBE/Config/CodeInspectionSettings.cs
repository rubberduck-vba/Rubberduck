using System.Xml.Serialization;
using Rubberduck.Inspections;

namespace Rubberduck.Config
{
    [XmlType(AnonymousType = true)]
    public class CodeInspectionSettings
    {
        [XmlArrayItem("CodeInspection", IsNullable = false)]
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

    [XmlType(AnonymousType = true)]
    public class CodeInspectionSetting : IInspectionModel
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        public string Description { get; set; }

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
            this.Description = name;
            this.InspectionType = type;
            this.Severity = severity;
        }

        public CodeInspectionSetting(IInspectionModel inspection)
            : this(inspection.Description, inspection.InspectionType, inspection.Severity)
        { }
    }
}
