using Rubberduck.VBEditor.SafeComWrappers.Office.Enums;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VB.Enums;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Abstract
{
    public interface ICommandBars : ISafeComWrapper, IComCollection<ICommandBar>
    {
        ICommandBar Add(string name);
        ICommandBar Add(string name, CommandBarPosition position);
        ICommandBarControl FindControl(int id);
        ICommandBarControl FindControl(ControlType type, int id);
    }
}