using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IComProjectDeserializer
    {
        ComProject DeserializeProject(ReferenceInfo reference);
        bool SerializedVersionExists(ReferenceInfo reference);
    }
}