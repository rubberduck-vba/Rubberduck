using System.Runtime.InteropServices;
using System.Xml.Serialization;
using Rubberduck.Inspections;

namespace Rubberduck.Config
{
    [ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class CodeInspectionSettings
    {
        [XmlArrayItemAttribute("CodeInspection", IsNullable = false)]
        public CodeInspection[] CodeInspections { get; set; }

        public CodeInspectionSettings()
        {
            //default constructor requied for serialization
        }

        public CodeInspectionSettings(CodeInspection[] inspections)
        {
            this.CodeInspections = inspections;
        }
    }

    [ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class CodeInspection
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        public CodeInspectionSeverity Severity { get; set; }

        [XmlAttribute]
        public CodeInspectionType InspectionType { get; set; }

        public CodeInspection()
        {
            //default constructor required for serialization
        }

        public CodeInspection(string name, CodeInspectionType type, CodeInspectionSeverity severity)
        {
            this.Name = name;
            this.InspectionType = type;
            this.Severity = severity;
        }

        public CodeInspection(Inspections.IInspection inspection)
            : this(inspection.Name, inspection.InspectionType, inspection.Severity)
        { }
    }
}
