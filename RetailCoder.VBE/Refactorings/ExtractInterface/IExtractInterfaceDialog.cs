using System.Collections.Generic;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public interface IExtractInterfaceDialog : IDialogView
    {
        string InterfaceName { get; set; }
        IEnumerable<InterfaceMember> Members { get; set; }
        List<string> ComponentNames { get; set; }
    }
}
