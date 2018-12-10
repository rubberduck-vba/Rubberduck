extern alias Office_v8;
using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using MSO = Office_v8::Office;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.Office8
{
    public class CommandBars : SafeComWrapper<MSO.CommandBars>, ICommandBars, IEquatable<ICommandBars>
    {
        private readonly IVBE _vbe;

        public CommandBars(MSO.CommandBars target, IVBE vbe, bool rewrapping = false) 
            : base(target, rewrapping)
        {
            _vbe = vbe;
        }

        public ICommandBar Add(string name)
        {
            DeleteExistingCommandBar(name);
            return new CommandBar(IsWrappingNullReference ? null : Target.Add(name, Temporary: true), _vbe);
        }

        public ICommandBar Add(string name, CommandBarPosition position)
        {
            DeleteExistingCommandBar(name);
            return new CommandBar(IsWrappingNullReference ? null : Target.Add(name, position, Temporary: true), _vbe);
        }

        private void DeleteExistingCommandBar(string name)
        {            
            foreach (var commandBar in (IEnumerable<ICommandBar>) this)
            {
                using (commandBar)
                {
                    if (commandBar.Name == name)
                    {
                        commandBar.Delete();
                    }
                }
            }
        }


        public ICommandBarControl FindControl(int id)
        {
            return new CommandBarControl(IsWrappingNullReference ? null : Target.FindControl(Id: id), _vbe);
        }

        public ICommandBarControl FindControl(ControlType type, int id)
        {
            return new CommandBarControl(IsWrappingNullReference ? null : Target.FindControl(type, id), _vbe);
        }

        IEnumerator<ICommandBar> IEnumerable<ICommandBar>.GetEnumerator()
        {
            return new ComWrapperEnumerator<ICommandBar>(Target,
                    comObject => new CommandBar((MSO.CommandBar) comObject, _vbe));
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable<ICommandBar>)this).GetEnumerator();
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public ICommandBar this[object index] => new CommandBar(IsWrappingNullReference ? null : Target[index], _vbe);

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