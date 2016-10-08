using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class CodePanes : SafeComWrapper<Microsoft.Vbe.Interop.CodePanes>, ICodePanes
    {
        public CodePanes(Microsoft.Vbe.Interop.CodePanes target) 
            : base(target)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : Target.Count; }
        }

        public IVBE Parent
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.Parent); }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }

        public ICodePane Current 
        { 
            get { return new CodePane(IsWrappingNullReference ? null : Target.Current); }
            set { Target.Current = (Microsoft.Vbe.Interop.CodePane)value.Target;}
        }

        public ICodePane this[object index]
        {
            get { return new CodePane(Target.Item(index)); }
        }

        IEnumerator<ICodePane> IEnumerable<ICodePane>.GetEnumerator()
        {
            return new ComWrapperEnumerator<CodePane>(Target);
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
                    this[i].Release();
                }
                Marshal.ReleaseComObject(Target);
            }
        }

        public override bool Equals(ISafeComWrapper<Microsoft.Vbe.Interop.CodePanes> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ICodePanes other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.CodePanes>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}