using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IComProjectSerializationProvider
    {
        string Target { get; }
        void SerializeProject(ComProject project);
        ComProject DeserializeProject(ReferenceInfo reference);
        bool SerializedVersionExists(ReferenceInfo reference);
    }
}
