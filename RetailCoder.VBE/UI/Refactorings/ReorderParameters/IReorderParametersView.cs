using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    interface IReorderParametersView : IDialogView
    {
        Declaration Target { get; set; }
        List<Parameter> Parameters { get; set; }
        void InitializeParameterGrid();
    }
}
