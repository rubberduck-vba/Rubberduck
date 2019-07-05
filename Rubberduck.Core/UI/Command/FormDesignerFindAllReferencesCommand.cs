using System.Runtime.InteropServices;

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
        {
            _findAllReferences = findAllReferences;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _findAllReferences.CanExecute(parameter);
        }

        protected override void OnExecute(object parameter)
        {
            _findAllReferences.Execute(parameter);
        }
    }
}
