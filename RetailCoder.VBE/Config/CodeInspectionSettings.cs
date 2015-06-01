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
            CodeInspections = inspections;
        }
    }

    [XmlType(AnonymousType = true)]
    public class CodeInspectionSetting : IInspectionModel
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlIgnore]
        public string Description { get; set; } // not serialized because culture-dependent

        [XmlAttribute]
        public CodeInspectionSeverity Severity { get; set; }

        [XmlAttribute]
        public CodeInspectionType InspectionType { get; set; }

        public CodeInspectionSetting()
        {
            //default constructor required for serialization
        }

        public CodeInspectionSetting(string name, string description, CodeInspectionType type, CodeInspectionSeverity severity)
        {
            Name = name;
            Description = description;
            InspectionType = type;
            Severity = severity;
        }

        public CodeInspectionSetting(IInspectionModel inspection)
            : this(inspection.Name, inspection.Description, inspection.InspectionType, inspection.Severity)
        { }
    }
}
