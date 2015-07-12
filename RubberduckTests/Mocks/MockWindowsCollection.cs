using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.UI;

namespace RubberduckTests.Mocks
{
    /// <summary>
    /// Fake Windows collection to get around Moq's inability to deal with ref params.
    /// </summary>
    /// <remarks>
    /// The <see cref="Window"/> passed into MockWindowCollection's ctor will be returned from <see cref="CreateToolWindow"/>.
    /// </remarks>
    class MockWindowsCollection : Windows, ICollection<Window>
    {
        internal MockWindowsCollection()
            :this(new List<Window>{MockFactory.CreateWindowMock().Object})
        { }

        internal MockWindowsCollection(ICollection<Window> windows)
        {
            _windows = windows;
        }

        private readonly ICollection<Window> _windows;

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        [SuppressMessage("ReSharper", "RedundantAssignment")]
        public Window CreateToolWindow(AddIn AddInInst, string ProgId, string Caption, string GuidPosition, ref object DocObj)
        {
            DocObj = new _DockableWindowHost(); 
            var result = MockFactory.CreateWindowMock(Caption);
            result.Setup(m => m.VBE).Returns(VBE);
            result.Setup(m => m.Collection).Returns(this);

            return result.Object;
        }

        public Window CreateWindow(string caption)
        {
            var result = MockFactory.CreateWindowMock(caption);
            result.Setup(m => m.VBE).Returns(VBE);
            result.Setup(m => m.Collection).Returns(this);

            return result.Object;
        }

        public void Add(Window window)
        {
            _windows.Add(window);
        }

        public bool Remove(Window window)
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

        public bool Contains(Window window)
        {
            return _windows.Contains(window);
        }

        public void CopyTo(Window[] array, int arrayIndex)
        {
            _windows.CopyTo(array, arrayIndex);
        }

        IEnumerator<Window> IEnumerable<Window>.GetEnumerator()
        {
            return _windows.GetEnumerator();
        }

        public IEnumerator GetEnumerator()
        {
            return _windows.GetEnumerator();
        }

        public Window Item(object index)
        {
            if (index is ValueType)
            {
                return _windows.ElementAt((int) index);
            }

            return _windows.FirstOrDefault(window => window.Caption == index.ToString());
        }

        public Application Parent
        {
            get { return VBE; }
        }

        public VBE VBE { get; set; }
    }
}
