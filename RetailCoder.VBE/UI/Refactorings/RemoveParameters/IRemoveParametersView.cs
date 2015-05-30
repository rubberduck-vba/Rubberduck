using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactoring.RemoveParameterRefactoring;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    interface IRemoveParametersView : IDialogView
    {
        RemoveParameterRefactoring RemoveParams { get; set; }
        void InitializeParameterGrid();
    }
}
