using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Serialization;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

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

        public CodeInspectionSettings() : this(Enumerable.Empty<CodeInspectionSetting>(), new WhitelistedIdentifierSetting[] { }, true)
        {
        }

        public CodeInspectionSettings(IEnumerable<CodeInspectionSetting> inspections, WhitelistedIdentifierSetting[] whitelistedNames, bool runInspectionsOnParse)
        {
            CodeInspections = new HashSet<CodeInspectionSetting>(inspections);
            WhitelistedIdentifiers = whitelistedNames;
            RunInspectionsOnSuccessfulParse = runInspectionsOnParse;
        }

        public CodeInspectionSetting GetSetting<TInspection>() where TInspection : IInspection
        {
            return CodeInspections.FirstOrDefault(s => typeof(TInspection).Name.ToString(CultureInfo.InvariantCulture).Equals(s.Name))
                ?? GetSetting(typeof(TInspection));
        }

        public CodeInspectionSetting GetSetting(Type inspectionType)
        {
            try
            {
                var existing = CodeInspections.FirstOrDefault(s => inspectionType.ToString().Equals(s.Name));
                if (existing != null)
                {
                    return existing;
                }
                var proto = Convert.ChangeType(Activator.CreateInstance(inspectionType, new object[]{null}), inspectionType);
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
        private string _description;
        [XmlIgnore]
        public string Description
        {
            get => _description ?? (_description = InspectionsUI.ResourceManager.GetString(Name + "Name"));
            set => _description = value;
        }// not serialized because culture-dependent

        [XmlIgnore]
        public string LocalizedName => InspectionsUI.ResourceManager.GetString(Name + "Name", CultureInfo.CurrentUICulture); // not serialized because culture-dependent

        [XmlIgnore]
        public string AnnotationName => Name.Replace("Inspection", string.Empty);

        [XmlIgnore]
        public CodeInspectionSeverity DefaultSeverity { get; }

        [XmlAttribute]
        public CodeInspectionSeverity Severity { get; set; }

        [XmlIgnore]
        public string Meta => InspectionsUI.ResourceManager.GetString(Name + "Meta", CultureInfo.CurrentUICulture);

        [XmlIgnore]
        // ReSharper disable once UnusedMember.Global; used in string literal to define collection groupings
        public string TypeLabel => InspectionsUI.ResourceManager.GetString("CodeInspectionSettings_" + InspectionType, CultureInfo.CurrentUICulture);

        [XmlIgnore]
        public string SeverityLabel
        {
            get => InspectionsUI.ResourceManager.GetString("CodeInspectionSeverity_" + Severity, CultureInfo.CurrentUICulture);
            set
            {
                foreach (var severity in Enum.GetValues(typeof(CodeInspectionSeverity)))
                {
                    if (value == InspectionsUI.ResourceManager.GetString("CodeInspectionSeverity_" + severity, CultureInfo.CurrentUICulture))
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

        public CodeInspectionSetting(string name, CodeInspectionType type, CodeInspectionSeverity defaultSeverity = CodeInspectionSeverity.Warning)
            : this(name, string.Empty, type, defaultSeverity, defaultSeverity)
        { }

        public CodeInspectionSetting(string name, string description, CodeInspectionType type, CodeInspectionSeverity defaultSeverity = CodeInspectionSeverity.Warning, CodeInspectionSeverity severity = CodeInspectionSeverity.Warning)
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
            return obj is CodeInspectionSetting inspectionSetting &&
                   inspectionSetting.InspectionType == InspectionType &&
                   inspectionSetting.Name == Name &&
                   inspectionSetting.Severity == Severity;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Name?.GetHashCode() ?? 0;
                hashCode = (hashCode * 397) ^ (int)Severity;
                hashCode = (hashCode * 397) ^ (int)InspectionType;
                return hashCode;
            }
        }
    }
}
