using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.DisposableWrappers.Office.Core
{
    public class CommandBarControls : SafeComWrapper<Microsoft.Office.Core.CommandBarControls>, IEnumerable<CommandBarControl>
    {
        public CommandBarControls(Microsoft.Office.Core.CommandBarControls comObject) 
            : base(comObject)
        {
        }

        public CommandBarControl this[object index]
        {
            get { return new CommandBarControl(InvokeResult(() => ComObject[index])); }
        }

        public CommandBarControl Add(ControlType type)
        {
            return new CommandBarControl(InvokeResult(() => ComObject.Add(type, Temporary:true)));
        }

        public CommandBarControl Add(ControlType type, int before)
        {
            return new CommandBarControl(InvokeResult(() => ComObject.Add(type, Before:before, Temporary:true)));
        }

        IEnumerator<CommandBarControl> IEnumerable<CommandBarControl>.GetEnumerator()
        {
            return new ComWrapperEnumerator<CommandBarControl>(ComObject);
        }

        public int Count { get { return InvokeResult(() => ComObject.Count); } }
        public CommandBar Parent { get { return new CommandBar(InvokeResult(() => ComObject.Parent)); } }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<CommandBarControl>)this).GetEnumerator();
        }
    }
}