using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.Parsing.ComReflection
{
    public class ComDocumentation
    {
        public string Name { get; private set; }
        public string DocString { get; private set; }
        public string HelpFile { get; private set; }
        public int HelpContext { get; private set; }

        public ComDocumentation(ITypeLib typeLib, int index)
        {
            LoadDocumentation(typeLib, null, index);
        }

        public ComDocumentation(ITypeInfo info, int index)
        {
            LoadDocumentation(null, info, index);
        }

        private void LoadDocumentation(ITypeLib typeLib, ITypeInfo info, int index)
        {
            string name;
            string docString;
            int helpContext;
            string helpFile;

            if (info == null)
            {
                typeLib.GetDocumentation(index, out name, out docString, out helpContext, out helpFile);
            }
            else
            {
                info.GetDocumentation(index, out name, out docString, out helpContext, out helpFile);
            }

            //See http://chat.stackexchange.com/transcript/message/30119269#30119269
            Name = name;
            DocString = docString;
            HelpContext = helpContext;
            HelpFile = helpFile;            
        }
    }
}
