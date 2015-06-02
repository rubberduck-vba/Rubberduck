using System.Collections.Generic;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    interface IRemoveParametersView : IDialogView
    {
        List<Parameter> Parameters { get; set; }
        void InitializeParameterGrid();
    }
}
