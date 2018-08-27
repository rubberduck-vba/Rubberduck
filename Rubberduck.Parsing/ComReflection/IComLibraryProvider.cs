using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IComLibraryProvider
    {
        ITypeLib LoadTypeLibrary(string libraryPath);
    }
}
