using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using VBAIA = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VBA
{
    public class Windows : SafeComWrapper<VBAIA.Windows>, IWindows
    {
        public Windows(VBAIA.Windows windows)
            : base(windows)
        {
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IApplication Parent => new Application(IsWrappingNullReference ? null : Target.Parent);

        public IWindow this[object index] => new Window(IsWrappingNullReference ? null : Target.Item(index));


        private static readonly Dictionary<VBAIA.Window, object> _dockableHosts = new Dictionary<VBAIA.Window, object>();

        public ToolWindowInfo CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition)
        {
            if (IsWrappingNullReference) return new ToolWindowInfo(null, null);
            object control = null;
            var window = Target.CreateToolWindow((VBAIA.AddIn)addInInst.Target, progId, caption, guidPosition, ref control);
            _dockableHosts.Add(window, control);
            return new ToolWindowInfo(new Window(window), control);
        }

        public static void ReleaseDockableHosts()
        {
            foreach (var item in _dockableHosts)
            {
                item.Key.Close();
                dynamic host = item.Value;
                host.Release();
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference ? new List<IEnumerable>().GetEnumerator() : Target.GetEnumerator();
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return IsWrappingNullReference
                ? new ComWrapperEnumerator<IWindow>(null, o => new Window(null))
                : new ComWrapperEnumerator<IWindow>(Target, o => new Window((VBAIA.Window) o));
        }

        public override bool Equals(ISafeComWrapper<VBAIA.Windows> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IWindows other)
        {
            return Equals(other as SafeComWrapper<VBAIA.Windows>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}