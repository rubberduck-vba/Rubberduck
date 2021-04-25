using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command.MenuItems
{
    public interface IAppMenu
    {
        void Localize();
        void Initialize();
        Task EvaluateCanExecuteAsync(RubberduckParserState state, CancellationToken token);
    }
}
