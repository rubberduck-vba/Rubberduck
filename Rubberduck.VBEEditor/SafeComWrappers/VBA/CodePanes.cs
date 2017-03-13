using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class CodePanes : SafeComWrapper<VB.CodePanes>, ICodePanes
    {
        public CodePanes(VB.CodePanes target) 
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
            set { if (!IsWrappingNullReference) Target.Current = (VB.CodePane)value.Target; }
        }

        public ICodePane this[object index]
        {
            get { return new CodePane(IsWrappingNullReference ? null : Target.Item(index)); }
        }

        IEnumerator<ICodePane> IEnumerable<ICodePane>.GetEnumerator()
        {
            return IsWrappingNullReference
                ? new ComWrapperEnumerator<ICodePane>(null, o => new CodePane(null))
                : new ComWrapperEnumerator<ICodePane>(Target, o => new CodePane((VB.CodePane) o));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator) new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<ICodePane>) this).GetEnumerator();
        }

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        for (var i = 1; i <= Count; i++)
        //        {
        //            this[i].Release();
        //        }
        //        base.Release(final);
        //    }
        //}

        public override bool Equals(ISafeComWrapper<VB.CodePanes> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ICodePanes other)
        {
            return Equals(other as SafeComWrapper<VB.CodePanes>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}