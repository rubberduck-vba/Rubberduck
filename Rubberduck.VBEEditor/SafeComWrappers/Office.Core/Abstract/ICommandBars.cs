using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;

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