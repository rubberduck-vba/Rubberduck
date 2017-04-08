using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class Windows : SafeComWrapper<VB.Windows>, IWindows
    {
        public Windows(VB.Windows windows)
            : base(windows)
        {
        }

        public int Count
        {
            get { return Target.Count; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }

        public IApplication Parent
        {
            get
            {
                throw new NotImplementedException();
                 /*return new Application(IsWrappingNullReference ? null : Target.Parent);*/
            }
        }

        public IWindow this[object index]
        {
            get { return new Window(Target.Item(index)); }
        }

        public ToolWindowInfo CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition)
        {
            object control = null;
            var window = new Window(Target.CreateToolWindow((VB.AddIn)addInInst.Target, progId, caption, guidPosition, ref control));
            return new ToolWindowInfo(window, control);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return Target.GetEnumerator();
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IWindow>(Target, o => new Window((VB.Window)o));
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

        public override bool Equals(ISafeComWrapper<VB.Windows> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IWindows other)
        {
            return Equals(other as SafeComWrapper<VB.Windows>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}