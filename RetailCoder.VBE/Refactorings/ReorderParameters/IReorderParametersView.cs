using System.Collections.Generic;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public interface IReorderParametersView : IDialogView
    {
        List<Parameter> Parameters { get; set; }
        void InitializeParameterGrid();
    }
}
