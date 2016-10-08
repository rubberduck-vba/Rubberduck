using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class AddIns : SafeComWrapper<Microsoft.Vbe.Interop.Addins>, IAddIns
    {
        public AddIns(Microsoft.Vbe.Interop.Addins comObject) : 
            base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Count; }
        }

        public object Parent // todo: verify if this could be 'public Application Parent' instead
        {
            get { return IsWrappingNullReference ? null : ComObject.Parent; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }

        public IAddIn this[object index]
        {
            get { return new AddIn(IsWrappingNullReference ? null : ComObject.Item(index)); }
        }

        public void Update()
        {
            ComObject.Update();
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

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Addins> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject.Parent, Parent));
        }

        public bool Equals(IAddIns other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Addins>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComputeHashCode(Parent);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ComObject.GetEnumerator();
        }

        IEnumerator<IAddIn> IEnumerable<IAddIn>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IAddIn>(ComObject);
        }
    }
}