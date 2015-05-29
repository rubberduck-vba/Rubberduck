using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactoring.ReorderParametersRefactoring;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    interface IReorderParametersView : IDialogView
    {
        ReorderParametersRefactoring ReorderParams { get; set; }
        void InitializeParameterGrid();
    }
}
