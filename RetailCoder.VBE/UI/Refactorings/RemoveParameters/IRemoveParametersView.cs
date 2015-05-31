using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    interface IRemoveParametersView : IDialogView
    {
        List<Parameter> Parameters { get; set; }
        void InitializeParameterGrid();
    }
}
