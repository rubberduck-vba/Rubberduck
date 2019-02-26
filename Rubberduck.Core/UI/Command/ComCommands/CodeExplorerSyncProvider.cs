using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.ComCommands
{
    public class CodeExplorerSyncProvider : ICodeExplorerSyncProvider
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IVBEEvents _vbeEvents;

        public CodeExplorerSyncProvider(IVBE vbe, RubberduckParserState state, IVBEEvents vbeEvents)
        {
            _vbe = vbe;
            _state = state;
            _vbeEvents = vbeEvents;
        }

        public SyncCodeExplorerCommand GetSyncCommand(CodeExplorerViewModel explorer)
        {
            return new SyncCodeExplorerCommand(_vbe, _state, explorer, _vbeEvents);
        }
    }
}