using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI
{
    [ComVisible(false)]
    public class SelectedTestEventArgs : EventArgs
    {
        public SelectedTestEventArgs(IEnumerable<TestExplorerItem> items)
        {
            _selection = items.Select(item => item.GetTestMethod());
        }

        public SelectedTestEventArgs(TestExplorerItem item)
            : this(new[] { item })
        { }

        private readonly IEnumerable<TestMethod> _selection;
        public IEnumerable<TestMethod> Selection { get { return _selection; } }
    }
}