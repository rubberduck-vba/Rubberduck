namespace Rubberduck.Parsing.ComReflection
{
    public interface IComProjectSerializationProvider : IComProjectDeserializer
    {
        string Target { get; }
        void SerializeProject(ComProject project);
    }
}
