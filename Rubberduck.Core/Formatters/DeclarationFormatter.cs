using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;

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
            return string.Format(RubberduckUI.CodeExplorer_IExportable_DeclarationFormat,
                _declaration.ProjectName,
                _declaration.CustomFolder,
                _declaration.ComponentName,
                _declaration.DeclarationType,
                _declaration.Scope,
                _declaration.IdentifierName);
        }
    }
}
