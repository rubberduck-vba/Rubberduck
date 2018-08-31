using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that locates all references to the active form designer component.
    /// </summary>
    [ComVisible(false)]
    public class FormDesignerFindAllReferencesCommand : CommandBase
    {
        private readonly FindAllReferencesCommand _findAllReferences;

        public FormDesignerFindAllReferencesCommand(FindAllReferencesCommand findAllReferences)
             : base(LogManager.GetCurrentClassLogger())
        {
            _findAllReferences = findAllReferences;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _findAllReferences.CanExecute(parameter);
        }

        protected override void OnExecute(object parameter)
        {
            _findAllReferences.Execute(parameter);
        }
    }
}
