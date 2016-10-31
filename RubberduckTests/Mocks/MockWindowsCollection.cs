using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.UI;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    /// <summary>
    /// Fake Windows collection to get around Moq's inability to deal with ref params.
    /// </summary>
    /// <remarks>
    /// The <see cref="Window"/> passed into MockWindowCollection's ctor will be returned from <see cref="CreateToolWindow"/>.
    /// </remarks>
    class MockWindowsCollection : IWindows, ICollection<IWindow>
    {
        internal MockWindowsCollection()
            :this(new List<IWindow>{MockFactory.CreateWindowMock().Object})
        { }

        internal MockWindowsCollection(ICollection<IWindow> windows)
        {
            _windows = windows;
        }

        private readonly ICollection<IWindow> _windows;

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        [SuppressMessage("ReSharper", "RedundantAssignment")]
        public IWindow CreateToolWindow(IAddIn AddInInst, string ProgId, string Caption, string GuidPosition, ref object DocObj)
        {
            DocObj = new _DockableWindowHost(); 
            var result = MockFactory.CreateWindowMock(Caption);
            result.Setup(m => m.VBE).Returns(VBE);
            result.Setup(m => m.Collection).Returns(this);

            return result.Object;
        }

        public IWindow CreateWindow(string caption)
        {
            var result = MockFactory.CreateWindowMock(caption);
            result.Setup(m => m.VBE).Returns(VBE);
            result.Setup(m => m.Collection).Returns(this);

            return result.Object;
        }

        public void Add(IWindow window)
        {
            _windows.Add(window);
        }

        public bool Remove(IWindow window)
        {
            return _windows.Remove(window);
        }

        public void Clear()
        {
            _windows.Clear();
        }

        public int Count
        {
            get { return _windows.Count; }
        }

        public bool IsReadOnly
        {
            get { return _windows.IsReadOnly; }
        }

        public bool Contains(IWindow window)
        {
            return _windows.Contains(window);
        }

        public void CopyTo(IWindow[] array, int arrayIndex)
        {
            _windows.CopyTo(array, arrayIndex);
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return _windows.GetEnumerator();
        }

        public IEnumerator GetEnumerator()
        {
            return _windows.GetEnumerator();
        }

        public IWindow this[object index]
        {
            get
            {
                if (index is ValueType)
                {
                    return _windows.ElementAt((int) index);
                }

                return _windows.FirstOrDefault(window => window.Caption == index.ToString());
            }
        }

        public IApplication Parent
        {
            get { throw new NotImplementedException(); }
        }

        public ToolWindowInfo CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition)
        {
            throw new NotImplementedException();
        }

        public IVBE VBE { get; set; }
        public object Target { get { throw new NotImplementedException(); } }

        public bool IsWrappingNullReference { get { throw new NotImplementedException(); } }

        public void Release(bool final = false)
        {
            throw new NotImplementedException();
        }

        public bool Equals(IWindows other)
        {
            throw new NotImplementedException();
        }
    }
}
