using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers.Office.Core
{
    public class CommandBarControls : SafeComWrapper<Microsoft.Office.Core.CommandBarControls>, IEnumerable<CommandBarControl>, IEquatable<CommandBarControls>
    {
        public CommandBarControls(Microsoft.Office.Core.CommandBarControls comObject) 
            : base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Count); }
        }

        public CommandBar Parent
        {
            get { return new CommandBar(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Parent)); }
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

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<CommandBarControl>)this).GetEnumerator();
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    this[i].Release();
                }
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Office.Core.CommandBarControls> other)
        {
            return IsEqualIfNull(other) || ReferenceEquals(other.ComObject, ComObject);
        }

        public bool Equals(CommandBarControls other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Office.Core.CommandBarControls>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}