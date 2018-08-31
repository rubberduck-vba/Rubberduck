using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.ComReflection
{
    public interface ISerializableProjectBuilder
    {
        SerializableProject SerializableProject(ProjectDeclaration projectDeclaration);
    }
}
