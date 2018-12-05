using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public interface ITokenStreamCache
    {
        ITokenStream CodePaneTokenStream(QualifiedModuleName module);
        ITokenStream AttributesTokenStream(QualifiedModuleName module);
    }
}