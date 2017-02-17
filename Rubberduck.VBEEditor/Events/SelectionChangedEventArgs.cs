using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class SelectionChangedEventArgs : EventArgs
    {
        public ICodePane CodePane { get; private set; }

        public SelectionChangedEventArgs(ICodePane pane)
        {
            CodePane = pane;
        }
    }
}
