using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using IInspection = Rubberduck.Parsing.Symbols.IInspection;

namespace Rubberduck.Settings
{
    public interface ICodeInspectionSettings
    {
        HashSet<CodeInspectionSetting> CodeInspections { get; set; }
        WhitelistedIdentifierSetting[] WhitelistedIdentifiers { get; set; }
        bool RunInspectionsOnSuccessfulParse { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class CodeInspectionSettings : ICodeInspectionSettings
    {
        [XmlArrayItem("CodeInspection", IsNullable = false)]
        public HashSet<CodeInspectionSetting> CodeInspections { get; set; }

        [XmlArrayItem("WhitelistedIdentifier", IsNullable = false)]
        public WhitelistedIdentifierSetting[] WhitelistedIdentifiers { get; set; }

        public bool RunInspectionsOnSuccessfulParse { get; set; }

        public CodeInspectionSettings() : this(new HashSet<CodeInspectionSetting>(), new WhitelistedIdentifierSetting[] {}, true)
        {
        }

        public CodeInspectionSettings(HashSet<CodeInspectionSetting> inspections, WhitelistedIdentifierSetting[] whitelistedNames, bool runInspectionsOnParse)
        {
            CodeInspections = inspections;
            WhitelistedIdentifiers = whitelistedNames;
            RunInspectionsOnSuccessfulParse = runInspectionsOnParse;
        }

        public CodeInspectionSetting GetSetting<TInspection>() where TInspection : IInspection
        {
            return CodeInspections.FirstOrDefault(s => typeof(TInspection).Name.ToString().Equals(s.Name))
                ?? GetSetting(typeof(TInspection));
        }

        public CodeInspectionSetting GetSetting(Type inspection)
        {
            try
            {
                var proto = Convert.ChangeType(Activator.CreateInstance(inspection), inspection);
                var existing = CodeInspections.FirstOrDefault(s => proto.GetType().ToString().Equals(s.Name));
                if (existing != null) return existing;
                var setting = new CodeInspectionSetting(proto as IInspection);
                CodeInspections.Add(setting);
                return setting;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }

    [XmlType(AnonymousType = true)]
    public class CodeInspectionSetting : IInspectionModel
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlIgnore]
        public CodeInspectionSeverity DefaultSeverity { get; private set; }

        [XmlAttribute]
        public CodeInspectionSeverity Severity { get; set; }

        [XmlAttribute]
        public CodeInspectionType InspectionType { get; set; }

        public CodeInspectionSetting()
        {
            //default constructor required for serialization
        }

        public CodeInspectionSetting(string name, CodeInspectionType type, CodeInspectionSeverity defaultSeverity, CodeInspectionSeverity severity)
        {
            Name = name;
            InspectionType = type;
            Severity = severity;
            DefaultSeverity = defaultSeverity;
        }

        public CodeInspectionSetting(IInspection inspection)
            : this(inspection.Name, inspection.InspectionType, inspection.DefaultSeverity, inspection.Severity)
        { }

        public override bool Equals(object obj)
        {
            var inspectionSetting = obj as CodeInspectionSetting;

            return inspectionSetting != null &&
                   inspectionSetting.InspectionType == InspectionType &&
                   inspectionSetting.Name == Name &&
                   inspectionSetting.Severity == Severity;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (Name != null ? Name.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (int) Severity;
                hashCode = (hashCode * 397) ^ (int) InspectionType;
                return hashCode;
            }
        }
    }
}
