using Microsoft.Win32;

namespace Rubberduck.Deployment.Build.Structs
{
    public struct RegistryEntry
    {
        public string Key { get; }
        public string Name { get; }
        public string Value { get; }
        public RegistryValueKind Type { get; }
        public Bitness Bitness { get; }

        public RegistryEntry(string key, string name, string value, RegistryValueKind type, Bitness bitness, FileMap fileMap)
        {
            Key = key;
            Name = name;
            Value = fileMap.Replace(value);
            Type = type;
            Bitness = bitness;
        }
    }
}