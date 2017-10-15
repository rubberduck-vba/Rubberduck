using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBars : SafeComWrapper<Microsoft.Office.Core.CommandBars>, ICommandBars, IEquatable<ICommandBars>
    {
        public CommandBars(Microsoft.Office.Core.CommandBars target) 
            : base(target)
        {
        }

        public ICommandBar Add(string name)
        {
            DeleteExistingCommandBar(name);
            return new CommandBar(IsWrappingNullReference ? null : Target.Add(name, Temporary: true));
        }

        public ICommandBar Add(string name, CommandBarPosition position)
        {
            DeleteExistingCommandBar(name);
            return new CommandBar(IsWrappingNullReference ? null : Target.Add(name, position, Temporary: true));
        }

        private void DeleteExistingCommandBar(string name)
        {
            try
            {
                var existing = Target.Cast<Microsoft.Office.Core.CommandBar>().FirstOrDefault(bar => bar.Name == name);
                if (existing != null)
                {
                    existing.Delete();
                    Marshal.FinalReleaseComObject(existing);
                }
            }
            catch
            {
                // specified commandbar didn't exist
            }
        }

        public ICommandBarControl FindControl(int id)
        {
            return new CommandBarControl(IsWrappingNullReference ? null : Target.FindControl(Id: id));
        }

        public ICommandBarControl FindControl(ControlType type, int id)
        {
            return new CommandBarControl(IsWrappingNullReference ? null : Target.FindControl(type, id));
        }

        IEnumerator<ICommandBar> IEnumerable<ICommandBar>.GetEnumerator()
        {
            return IsWrappingNullReference
                ? new ComWrapperEnumerator<ICommandBar>(null, o => new CommandBar(null))
                : new ComWrapperEnumerator<ICommandBar>(Target,
                    o => new CommandBar((Microsoft.Office.Core.CommandBar) o));
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable<ICommandBar>)this).GetEnumerator();
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : Target.Count; }
        }

        public ICommandBar this[object index]
        {
            get { return new CommandBar(IsWrappingNullReference ? null : Target[index]); }
        }

        //public override void Release(bool final = false)
        //{
        //    // important: no-op
        //}

        public override bool Equals(ISafeComWrapper<Microsoft.Office.Core.CommandBars> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ICommandBars other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Office.Core.CommandBars>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}