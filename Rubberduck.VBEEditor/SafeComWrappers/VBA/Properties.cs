using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Properties : SafeComWrapper<Microsoft.Vbe.Interop.Properties>, IEnumerable<Property>, IEquatable<Properties>
    {
        public Properties(Microsoft.Vbe.Interop.Properties comObject) 
            : base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Count; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }

        public IApplication Application
        {
            get { return new Application(IsWrappingNullReference ? null : ComObject.Application); }
        }

        public object Parent
        {
            get { return IsWrappingNullReference ? null : ComObject.Parent; }
        }

        public Property Item(object index)
        {
            return new Property(ComObject.Item(index));
        }

        IEnumerator<Property> IEnumerable<Property>.GetEnumerator()
        {
            return new ComWrapperEnumerator<Property>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<Property>)this).GetEnumerator();
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    Item(i).Release();
                }
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Properties> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(Properties other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Properties>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}