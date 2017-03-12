using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Moq;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    public class Windows : SafeComWrapper<object>, IWindows, ICollection<IWindow>
    {
        private readonly IList<IWindow> _windows = new List<IWindow>();
        
        public Windows()
            : base(new object())
        {
        }

        public void Add(IWindow item)
        {
            _windows.Add(item);
        }

        public void Clear()
        {
            _windows.Clear();
        }

        public bool Contains(IWindow item)
        {
            return _windows.Contains(item);
        }

        public void CopyTo(IWindow[] array, int arrayIndex)
        {
            _windows.CopyTo(array, arrayIndex);
        }

        public bool Remove(IWindow item)
        {
            return _windows.Remove(item);
        }

        public int Count
        {
            get { return _windows.Count; }
        }

        public bool IsReadOnly { get { return _windows.IsReadOnly; } }

        public IVBE VBE { get; set; }

        public IApplication Parent
        {
            get { return null; }
        }

        public IWindow this[object index]
        {
            get
            {
                if (index is string)
                {
                    return _windows.SingleOrDefault(window => window.Caption == index.ToString());
                }

                return _windows[(int)index];
            }
        }

        public ToolWindowInfo CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition)
        {
            var window = new Mock<IWindow>();
            return new ToolWindowInfo(window.Object, null);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _windows.GetEnumerator();
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return _windows.GetEnumerator();
        }

        //public override void Release(bool final = false)
        //{
        //}

        public override bool Equals(ISafeComWrapper<object> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IWindows other)
        {
            return Equals(other as SafeComWrapper<object>);
        }

        public override int GetHashCode()
        {
            return _windows.GetHashCode();
        }

        public IWindow CreateWindow(string name)
        {
            var result = new Mock<IWindow>();
            result.Setup(m => m.Caption).Returns(name);
            return result.Object;
        }
    }
}