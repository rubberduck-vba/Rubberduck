using System;
using System.Linq;
using System.Runtime.Serialization;

namespace Rubberduck.Parsing.ComReflection
{
    [DataContract]
    [KnownType(typeof(ComProject))]
    public class ComTypeName
    {
        [DataMember(IsRequired = true)]
        public Guid EnumGuid { get; private set; } = Guid.Empty;
        public bool IsEnumMember => !EnumGuid.Equals(Guid.Empty);

        [DataMember(IsRequired = true)]
        public Guid AliasGuid { get; private set; } = Guid.Empty;
        public bool IsAliased => !AliasGuid.Equals(Guid.Empty);

        public ComProject Project { get; set; }

        [DataMember(IsRequired = true)]
        private string _rawName;
        public string Name
        {
            get
            {
                if (IsEnumMember && ComProject.KnownEnumerations.TryGetValue(EnumGuid, out var enumeration))
                {
                    return enumeration.Name;
                }

                if (IsAliased && ComProject.KnownAliases.TryGetValue(AliasGuid, out var alias))
                {
                    return alias.Name;
                }

                if (Project == null)
                {
                    return _rawName;
                }

                var softAlias = Project.Aliases.FirstOrDefault(x => x.Name.Equals(_rawName));
                return softAlias == null ? _rawName : softAlias.TypeName;
            }
        }

        public ComTypeName(ComProject project, string name)
        {
            Project = project;
            _rawName = name;
        }

        public ComTypeName(ComProject project, string name, Guid enumGuid, Guid aliasGuid) : this(project, name)
        {
            EnumGuid = enumGuid;
            AliasGuid = aliasGuid;
        }
    }
}
