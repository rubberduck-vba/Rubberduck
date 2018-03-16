using System.Collections.Generic;

namespace Rubberduck.Deployment.Structs
{
    public struct RecordMap
    {
        public string Guid { get; }
        public List<RegistryEntry> Entries { get; }

        public RecordMap(string guid, List<RegistryEntry> entries)
        {
            Guid = guid;
            Entries = entries;
        }
    }
}