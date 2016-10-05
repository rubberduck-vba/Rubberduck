using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class CodePanes : SafeComWrapper<Microsoft.Vbe.Interop.CodePanes>, IEnumerable<CodePane>, IEquatable<CodePane>
    {
        public CodePanes(Microsoft.Vbe.Interop.CodePanes comObject) 
            : base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Count); }
        }

        public VBE Parent
        {
            get { return new VBE(InvokeResult(() => IsWrappingNullReference ? null : ComObject.Parent)); }
        }

        public VBE VBE
        {
            get { return new VBE(InvokeResult(() => IsWrappingNullReference ? null : ComObject.VBE)); }
        }

        public CodePane Current 
        { 
            get { return new CodePane(InvokeResult(() => ComObject.Current)); }
            set { Invoke(() => ComObject.Current = value.ComObject);}
        }

        public CodePane Item(object index)
        {
            return new CodePane(InvokeResult(() => ComObject.Item(index)));
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
            return IsEqualIfNull(other) || ReferenceEquals(other.ComObject, ComObject);
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