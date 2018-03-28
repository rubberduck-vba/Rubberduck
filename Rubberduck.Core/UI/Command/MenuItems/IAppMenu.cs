using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command.MenuItems
{
    public interface IAppMenu
    {
        void Localize();
        void Initialize();
        void EvaluateCanExecute(RubberduckParserState state);
    }
}
