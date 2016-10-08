using Rubberduck.VBEditor.SafeComWrappers.Forms;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract
{
    public interface ICommandBars : ISafeComWrapper, IComCollection<ICommandBar>
    {
        ICommandBar Add(string name);
        ICommandBar Add(string name, CommandBarPosition position);
        ICommandBarControl FindControl(int id);
        ICommandBarControl FindControl(ControlType type, int id);
    }
}