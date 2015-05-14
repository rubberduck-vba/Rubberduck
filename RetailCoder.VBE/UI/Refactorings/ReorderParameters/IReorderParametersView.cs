using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    interface IReorderParametersView : IDialogView
    {
        Declaration Target { get; set; }
    }
}
