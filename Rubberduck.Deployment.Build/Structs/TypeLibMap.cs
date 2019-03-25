using System.Collections.Generic;

namespace Rubberduck.Deployment.Build.Structs
{
    public struct TypeLibMap
    {
        public string Guid { get; }
        public string Version { get; }
        public List<RegistryEntry> Entries { get; }

        public TypeLibMap(string guid, string version, List<RegistryEntry> entries)
        {
            Guid = guid;
            Version = version;
            Entries = entries;
        }
    }
}