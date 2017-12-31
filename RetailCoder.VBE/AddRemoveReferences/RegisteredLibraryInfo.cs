using System;

namespace Rubberduck.AddRemoveReferences
{
    public struct RegisteredLibraryInfo
    {
        public string Name { get; set; }
        public string Guid { get; set; }
        public string Description { get; set; }
        public Version Version { get; set; }
        public string FullPath { get; set; }
    }
}