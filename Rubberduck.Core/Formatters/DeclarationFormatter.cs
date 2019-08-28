using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Formatters
{
    public class DeclarationFormatter : IExportable
    {
        private readonly Declaration _declaration;

        public DeclarationFormatter(Declaration declaration)
        {
            _declaration = declaration;
        }

        public object[] ToArray()
        {
            return _declaration.ToArray();
        }

        public string ToClipboardString()
        {
            return _declaration.ToString(); //TODO: Needs proper formatting
        }
    }
}
