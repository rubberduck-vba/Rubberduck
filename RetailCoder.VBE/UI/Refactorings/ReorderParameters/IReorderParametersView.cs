using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    interface IReorderParametersView : IDialogView
    {
        Declaration Target { get; set; }
        List<Parameter> Parameters { get; set; }
        void InitializeParameterGrid();
    }
}
