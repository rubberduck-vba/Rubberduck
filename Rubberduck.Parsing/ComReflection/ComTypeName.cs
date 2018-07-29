using System;
using System.Linq;

namespace Rubberduck.Parsing.ComReflection
{
    public class ComTypeName
    {
        public Guid EnumGuid { get; } = Guid.Empty;
        public bool IsEnumMember => !EnumGuid.Equals(Guid.Empty);

        public Guid AliasGuid { get; } = Guid.Empty;
        public bool IsAliased => !AliasGuid.Equals(Guid.Empty);

        public ComProject Project { get; }

        private readonly string _rawName;
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
