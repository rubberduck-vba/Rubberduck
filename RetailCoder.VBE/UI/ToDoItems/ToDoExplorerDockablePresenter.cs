using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.ToDoItems
{
    /// <summary>
    /// Presenter for the to-do items explorer.
    /// </summary>
    public class ToDoExplorerDockablePresenter : DockableToolwindowPresenter
    {

        public ToDoExplorerDockablePresenter(VBE vbe, AddIn addin, IDockableUserControl window)
            : base(vbe, addin, window)
        {
        }
    }
}
