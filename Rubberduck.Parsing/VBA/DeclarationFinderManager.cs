using Rubberduck.VBEditor.Application;
using System;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public class DeclarationFinderManager : IDeclarationFinderManager 
    {
        private readonly RubberduckParserState _state;
        private readonly IHostApplication _hostApp;

        public DeclarationFinderManager(
            RubberduckParserState state,
            IVBE vbe)
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            if (vbe == null) throw new ArgumentNullException(nameof(vbe));

            _state = state;
            _hostApp = vbe.HostApplication();
        }

        public DeclarationFinder DeclarationFinder
        {
            get
            {
                return _state.DeclarationFinder;
            }
        }

        public void RefreshDeclarationFinder()
        {
            _state.RefreshFinder(_hostApp);
        }
    }
}
