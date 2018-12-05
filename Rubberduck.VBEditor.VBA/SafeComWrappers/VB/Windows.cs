using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public sealed class Windows : SafeComWrapper<VB.Windows>, IWindows
    {
        public Windows(VB.Windows target, bool rewrapping = false)
            : base(target, rewrapping)
        {
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IApplication Parent => new Application(IsWrappingNullReference ? null : Target.Parent);

        public IWindow this[object index] => new Window(IsWrappingNullReference ? null : Target.Item(index));


        private static readonly Dictionary<VB.Window, object> _dockableHosts = new Dictionary<VB.Window, object>();

        public ToolWindowInfo CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition)
        {
            if (IsWrappingNullReference)
            {
                return new ToolWindowInfo(null, null);
            }
            object control = null;
            var window = Target.CreateToolWindow((VB.AddIn)addInInst.Target, progId, caption, guidPosition, ref control);
            _dockableHosts.Add(window, control);
            return new ToolWindowInfo(new Window(window), control);
        }

        public void ReleaseDockableHosts()
        {
            foreach (var item in _dockableHosts)
            {
                item.Key.Close();
                dynamic host = item.Value;
                host.Release();
            }
            _dockableHosts.Clear();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference ? new List<IEnumerable>().GetEnumerator() : Target.GetEnumerator();
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IWindow>(Target, comObject => new Window((VB.Window) comObject));
        }

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

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}