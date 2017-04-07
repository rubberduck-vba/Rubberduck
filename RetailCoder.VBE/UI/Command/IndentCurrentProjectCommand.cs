using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentProjectCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly IIndenter _indenter;
        private readonly RubberduckParserState _state;

        public IndentCurrentProjectCommand(IVBE vbe, IIndenter indenter, RubberduckParserState state) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _indenter = indenter;
            _state = state;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return !_vbe.ActiveVBProject.IsWrappingNullReference && _vbe.ActiveVBProject.Protection != ProjectProtection.Locked;
        }

        protected override void ExecuteImpl(object parameter)
        {
            _indenter.IndentCurrentProject();
            _state.OnParseRequested(this);
        }
    }
}
