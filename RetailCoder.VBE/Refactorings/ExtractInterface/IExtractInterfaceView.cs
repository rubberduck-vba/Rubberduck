using System.Collections.Generic;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public interface IExtractInterfaceView : IDialogView
    {
        string InterfaceName { get; set; }
        List<InterfaceMember> Members { get; set; }
        List<string> ComponentNames { get; set; }
    }
}