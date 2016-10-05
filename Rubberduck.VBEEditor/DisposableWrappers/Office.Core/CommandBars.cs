using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers.Office.Core
{
    public class CommandBars : SafeComWrapper<Microsoft.Office.Core.CommandBars>, IEnumerable<CommandBar>, IEquatable<CommandBars>
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

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Count); }
        }

        public CommandBar this[object index]
        {
            get { return new CommandBar(InvokeResult(() => ComObject[index])); }
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                var commandBars = this.ToArray();
                foreach (var commandBar in commandBars)
                {
                    commandBar.Release();
                }
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Office.Core.CommandBars> other)
        {
            return IsEqualIfNull(other) || ReferenceEquals(other.ComObject, ComObject);
        }

        public bool Equals(CommandBars other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Office.Core.CommandBars>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}