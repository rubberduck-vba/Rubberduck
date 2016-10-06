using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class CodePanes : SafeComWrapper<Microsoft.Vbe.Interop.CodePanes>, IEnumerable<CodePane>, IEquatable<CodePane>
    {
        public CodePanes(Microsoft.Vbe.Interop.CodePanes comObject) 
            : base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Count; }
        }

        public IVBE Parent
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.Parent); }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }

        public CodePane Current 
        { 
            get { return new CodePane(IsWrappingNullReference ? null : ComObject.Current); }
            set { ComObject.Current = value.ComObject;}
        }

        public CodePane Item(object index)
        {
            return new CodePane(ComObject.Item(index));
        }

        IEnumerator<CodePane> IEnumerable<CodePane>.GetEnumerator()
        {
            return new ComWrapperEnumerator<CodePane>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<CodePane>)this).GetEnumerator();
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

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.CodePanes> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(CodePane other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.CodePane>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}