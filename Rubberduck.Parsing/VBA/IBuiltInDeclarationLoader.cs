
namespace Rubberduck.Parsing.VBA
{
    public interface IBuiltInDeclarationLoader
    {
        bool LastExecutionLoadedDeclarations { get; }

        void LoadBuitInDeclarations();
    }
}
