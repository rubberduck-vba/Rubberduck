using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.Parsing.ComReflection
{
    public class ComDocumentation
    {
        public const int LibraryIndex = -1;

        public string Name { get; }
        public string DocString { get; }
        public string HelpFile { get; }
        public int HelpContext { get; }

        public ComDocumentation(ITypeLib typeLib, int index)
        {
            typeLib.GetDocumentation(index, out string name, out string docString, out int helpContext, out string helpFile);
            Name = name;
            DocString = docString;
            HelpContext = helpContext;
            HelpFile = helpFile;
        }

        public ComDocumentation(ITypeInfo info, int index)
        {
            info.GetDocumentation(index, out string name, out string docString, out int helpContext, out string helpFile);
            Name = name;
            DocString = docString;
            HelpContext = helpContext;
            HelpFile = helpFile;
        }
    }
}
