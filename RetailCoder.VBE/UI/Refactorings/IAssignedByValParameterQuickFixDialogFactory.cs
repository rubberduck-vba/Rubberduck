using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.Refactorings
{
    public interface IAssignedByValParameterQuickFixDialogFactory
    {
        IAssignedByValParameterQuickFixDialog Create(string identifier, string identifierType, IEnumerable<string> forbiddenNames);
    }
}
