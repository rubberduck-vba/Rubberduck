using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using MSO = Microsoft.Office.Core;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.Office12
{
    public class CommandBars : SafeComWrapper<MSO.CommandBars>, ICommandBars, IEquatable<ICommandBars>
    {
        public CommandBars(MSO.CommandBars target, bool rewrapping = false) 
            : base(target, rewrapping)
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
                var existing = Target.Cast<MSO.CommandBar>().FirstOrDefault(bar => bar.Name == name);
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
            return new ComWrapperEnumerator<ICommandBar>(Target,
                    comObject => new CommandBar((MSO.CommandBar) comObject));
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable<ICommandBar>)this).GetEnumerator();
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public ICommandBar this[object index] => new CommandBar(IsWrappingNullReference ? null : Target[index]);

        public override bool Equals(ISafeComWrapper<MSO.CommandBars> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ICommandBars other)
        {
            return Equals(other as SafeComWrapper<MSO.CommandBars>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}