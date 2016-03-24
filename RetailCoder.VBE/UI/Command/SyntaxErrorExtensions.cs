using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command
{
    public static class SyntaxErrorExtensions
    {
        public static NavigateCodeEventArgs GetNavigateCodeEventArgs(this SyntaxErrorException exception, Declaration declaration)
        {
            var selection = new Selection(exception.LineNumber, exception.Position, exception.LineNumber, exception.Position);
            return new NavigateCodeEventArgs(declaration.QualifiedName.QualifiedModuleName, selection);
        }
    }
}