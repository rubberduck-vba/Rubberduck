
namespace Rubberduck.Parsing.VBA
{
    public interface IBuiltInDeclarationLoader
    {
        bool LastLoadOfBuiltInDeclarationsLoadedDeclarations { get; }

        void LoadBuitInDeclarations();
    }
}
