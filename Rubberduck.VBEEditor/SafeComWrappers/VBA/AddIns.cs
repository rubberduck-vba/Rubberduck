using System;
using System.Collections;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class AddIns : SafeComWrapper<Microsoft.Vbe.Interop.Addins>, IEnumerable, IEquatable<AddIns>
    {
        public AddIns(Microsoft.Vbe.Interop.Addins comObject) : 
            base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Count); }
        }

        public object Parent // todo: verify if this could be 'public Application Parent' instead
        {
            get { return IsWrappingNullReference ? null : InvokeResult(() => ComObject.Parent); }
        }

        public VBE VBE
        {
            get { return IsWrappingNullReference ? null : new VBE(InvokeResult(() => ComObject.VBE)); }
        }

        public AddIn Item(object index)
        {
            return new AddIn(InvokeResult(() => ComObject.Item(index)));
        }

        public void Update()
        {
            Invoke(() => ComObject.Update());
        }

        public IEnumerator GetEnumerator()
        {
            return InvokeResult(() => ComObject.GetEnumerator());
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

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Addins> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject.Parent, Parent));
        }

        public bool Equals(AddIns other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Addins>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComputeHashCode(Parent);
        }
    }
}