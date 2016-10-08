using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Forms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBars : SafeComWrapper<Microsoft.Office.Core.CommandBars>, ICommandBars, IEquatable<ICommandBars>
    {
        public CommandBars(Microsoft.Office.Core.CommandBars comObject) 
            : base(comObject)
        {
        }

        public ICommandBar Add(string name)
        {
            return new CommandBar(ComObject.Add(name, Temporary:true));
        }

        public ICommandBar Add(string name, CommandBarPosition position)
        {
            return new CommandBar(ComObject.Add(name, position, Temporary: true));
        }

        public ICommandBarControl FindControl(int id)
        {
            return new CommandBarControl(ComObject.FindControl(Id:id));
        }

        public ICommandBarControl FindControl(ControlType type, int id)
        {
            return new CommandBarControl(ComObject.FindControl(type, id));
        }

        IEnumerator<ICommandBar> IEnumerable<ICommandBar>.GetEnumerator()
        {
            return new ComWrapperEnumerator<CommandBar>(ComObject);
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable<ICommandBar>)this).GetEnumerator();
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Count; }
        }

        public ICommandBar this[object index]
        {
            get { return new CommandBar(ComObject[index]); }
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
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(ICommandBars other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Office.Core.CommandBars>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}