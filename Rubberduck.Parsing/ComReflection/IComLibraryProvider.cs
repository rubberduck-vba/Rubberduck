using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IComLibraryProvider
    {
        ITypeLib LoadTypeLibrary(string libraryPath);
        IComDocumentation GetComDocumentation(ITypeLib typelib);
        ReferenceInfo GetReferenceInfo(ITypeLib typelib, string name, string path);
    }
}
