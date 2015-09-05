using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that runs all active code inspections for the active VBAProject.
    /// </summary>
    [ComVisible(false)]
    public class RunCodeInspectionsCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly IInspector _inspector;
        private readonly IRubberduckParser _parser;

        public RunCodeInspectionsCommand(VBE vbe, IInspector inspector, IRubberduckParser parser)
        {
            _vbe = vbe;
            _inspector = inspector;
            _parser = parser;
        }

        /// <summary>
        /// Runs code inspections 
        /// </summary>
        /// <param name="parameter"></param>
        public override async void Execute(object parameter)
        {
            // todo: find a way to run this from UI
            //var project = parameter as VBProject;
            //var parseResult = project == null 
            //    ? _parser.Parse(_vbe.ActiveVBProject)
            //    : _parser.Parse(project);

            //var results = _inspector.FindIssuesAsync(parseResult, CancellationToken.None);
        }
    }

    public class RunCodeInspectionsCommandMenuItem : CommandMenuItemBase
    {
        public RunCodeInspectionsCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_CodeInspections"; } }
        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.CodeInspections; } }
    }
}