using System.Collections.Generic;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public interface IRemoveParametersDialog : IDialogView
    {
        List<Parameter> Parameters { get; set; }
        void InitializeParameterGrid();
    }
}
