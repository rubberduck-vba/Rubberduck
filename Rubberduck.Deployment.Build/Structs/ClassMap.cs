using System.Collections.Generic;

namespace Rubberduck.Deployment.Build.Structs
{
    public struct ClassMap
    {
        public string Guid { get; }
        public string Context { get; }
        public string Description { get; }
        public string ThreadingModel { get; }
        public string ProgId { get; }
        public string ProgIdDescription { get; }
        public List<RegistryEntry> Entries { get; }

        public ClassMap(string guid, string context, string description, string threadingModel, string progId,
            string progIdDescription, List<RegistryEntry> entries)
        {
            Guid = guid;
            Context = context;
            Description = description;
            ThreadingModel = threadingModel;
            ProgId = progId;
            ProgIdDescription = progIdDescription;
            Entries = entries;
        }
    }
}