using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Serialization;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.UI;

namespace Rubberduck.Settings
{
    public interface ICodeInspectionSettings
    {
        HashSet<CodeInspectionSetting> CodeInspections { get; set; }
        WhitelistedIdentifierSetting[] WhitelistedIdentifiers { get; set; }
        bool RunInspectionsOnSuccessfulParse { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class CodeInspectionSettings : ICodeInspectionSettings, IEquatable<CodeInspectionSettings>
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
            return CodeInspections.FirstOrDefault(s => typeof(TInspection).Name.ToString(CultureInfo.InvariantCulture).Equals(s.Name))
                ?? GetSetting(typeof(TInspection));
        }

        public CodeInspectionSetting GetSetting(Type inspection)
        {
            try
            {
                var proto = Convert.ChangeType(Activator.CreateInstance(inspection), inspection);
                var existing = CodeInspections.FirstOrDefault(s => proto.GetType().ToString().Equals(s.Name));
                if (existing != null) return existing;
                var setting = new CodeInspectionSetting(proto as IInspectionModel);
                CodeInspections.Add(setting);
                return setting;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public bool Equals(CodeInspectionSettings other)
        {
            return other != null &&
                   CodeInspections.SequenceEqual(other.CodeInspections) &&
                   WhitelistedIdentifiers.SequenceEqual(other.WhitelistedIdentifiers) &&
                   RunInspectionsOnSuccessfulParse == other.RunInspectionsOnSuccessfulParse;
        }
    }

    [XmlType(AnonymousType = true)]
    public class CodeInspectionSetting : IInspectionModel
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlIgnore]
        public string Description { get; set; } // not serialized because culture-dependent

        [XmlIgnore]
        public string LocalizedName
        {
            get
            {
                return InspectionsUI.ResourceManager.GetString(Name + "Name", CultureInfo.CurrentUICulture);
            }
        } // not serialized because culture-dependent

        [XmlIgnore]
        public string AnnotationName { get; set; }

        [XmlIgnore]
        public CodeInspectionSeverity DefaultSeverity { get; private set; }

        [XmlAttribute]
        public CodeInspectionSeverity Severity { get; set; }

        [XmlIgnore]
        public string Meta
        {
            get
            {
                return InspectionsUI.ResourceManager.GetString(Name + "Meta", CultureInfo.CurrentUICulture);
            }
        }

        [XmlIgnore]
        public string TypeLabel
        {
            get { return RubberduckUI.ResourceManager.GetString("CodeInspectionSettings_" + InspectionType, CultureInfo.CurrentUICulture); }
        }

        [XmlIgnore]
        public string SeverityLabel
        {
            get { return RubberduckUI.ResourceManager.GetString("CodeInspectionSeverity_" + Severity, CultureInfo.CurrentUICulture); }
            set
            {
                foreach (var severity in Enum.GetValues(typeof (CodeInspectionSeverity)))
                {
                    if (value == RubberduckUI.ResourceManager.GetString("CodeInspectionSeverity_" + severity, CultureInfo.CurrentUICulture))
                    {
                        Severity = (CodeInspectionSeverity)severity;
                        return;
                    }
                }
            }
        }

        [XmlAttribute]
        public CodeInspectionType InspectionType { get; set; }

        public CodeInspectionSetting()
        {
            //default constructor required for serialization
        }

        public CodeInspectionSetting(string name, string description, CodeInspectionType type, CodeInspectionSeverity defaultSeverity, CodeInspectionSeverity severity)
        {
            Name = name;
            Description = description;
            InspectionType = type;
            Severity = severity;
            DefaultSeverity = defaultSeverity;
        }

        public CodeInspectionSetting(IInspectionModel inspection)
            : this(inspection.Name, inspection.Description, inspection.InspectionType, inspection.DefaultSeverity, inspection.Severity)
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
