using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactoring;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    interface IReorderParametersView : IDialogView
    {
        List<Parameter> Parameters { get; set; }
        void InitializeParameterGrid();
    }
}
