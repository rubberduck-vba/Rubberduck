using Rubberduck.Navigation.CodeExplorer;

namespace Rubberduck.UI.Command.ComCommands
{
    public interface ICodeExplorerSyncProvider
    {
        SyncCodeExplorerCommand GetSyncCommand(CodeExplorerViewModel explorer);
    }
}