using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.DisposableWrappers.Office.Core
{
    public class CommandBars : SafeComWrapper<Microsoft.Office.Core.CommandBars>, IEnumerable<CommandBar>
    {
        public CommandBars(Microsoft.Office.Core.CommandBars comObject) 
            : base(comObject)
        {
        }

        public CommandBar Add(string name)
        {
            return new CommandBar(InvokeResult(() => ComObject.Add(name, Temporary:true)));
        }

        public CommandBar Add(string name, CommandBarPosition position)
        {
            return new CommandBar(InvokeResult(() => ComObject.Add(name, position, Temporary: true)));
        }

        public CommandBarControl FindControl(int id)
        {
            return new CommandBarControl(InvokeResult(() => ComObject.FindControl(Id:id)));
        }

        public CommandBarControl FindControl(ControlType type, int id)
        {
            return new CommandBarControl(InvokeResult(() => ComObject.FindControl(type, id)));
        }

        IEnumerator<CommandBar> IEnumerable<CommandBar>.GetEnumerator()
        {
            return new ComWrapperEnumerator<CommandBar>(ComObject);
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable<CommandBar>)this).GetEnumerator();
        }

        public int Count { get { return InvokeResult(() => ComObject.Count); } }

        public CommandBar this[object index]
        {
            get { return new CommandBar(InvokeResult(() => ComObject[index])); }
        }
    }
}