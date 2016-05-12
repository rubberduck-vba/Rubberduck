using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentModuleCommand : CommandBase
    {
        private readonly IIndenter _indenter;

        public IndentCurrentModuleCommand(IIndenter indenter)
        {
            _indenter = indenter;
        }

        public override bool CanExecute(object parameter)
        {
            var target = parameter as Declaration;
            return target != null && target.Annotations.All(a => a.AnnotationType != AnnotationType.NoIndent);
        }

        public override void Execute(object parameter)
        {
            _indenter.IndentCurrentModule();
        }

        public RubberduckHotkey Hotkey { get {return RubberduckHotkey.IndentModule; } }
    }
}