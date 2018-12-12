using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IComDocumentation
    {
        string Name { get; }
        string DocString { get; }
        string HelpFile { get; }
        int HelpContext { get; }
    }

    [DataContract]
    public class ComDocumentation : IComDocumentation
    {
        public const int LibraryIndex = -1;

        [DataMember(IsRequired = true)]
        public string Name { get; private set; }

        [DataMember(IsRequired = true)]
        public string DocString { get; private set; }

        [DataMember(IsRequired = true)]
        public string HelpFile { get; private set; }

        [DataMember(IsRequired = true)]
        public int HelpContext { get; private set; }

        public ComDocumentation(ITypeLib typeLib, int index)
        {
            typeLib.GetDocumentation(index, out string name, out string docString, out int helpContext, out string helpFile);
            Name = name;
            DocString = docString;
            HelpContext = helpContext;
            HelpFile = helpFile?.Trim('\0');
        }

        public ComDocumentation(ITypeInfo info, int index)
        {
            info.GetDocumentation(index, out string name, out string docString, out int helpContext, out string helpFile);
            Name = name;
            DocString = docString;
            HelpContext = helpContext;
            HelpFile = helpFile?.Trim('\0');
        }
    }
}
