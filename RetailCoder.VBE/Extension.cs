using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Extensions.Factory;
using Rubberduck.Root;
using Rubberduck.UI;

namespace Rubberduck
{
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgId)]
    [EditorBrowsable(EditorBrowsableState.Never)]
    // ReSharper disable once InconsistentNaming // note: underscore prefix hides class from COM API
    public class _Extension : IDTExtensibility2
    {
        private const string ClassId = "8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66";
        private const string ProgId = "Rubberduck.Extension";

        private readonly IKernel _kernel = new StandardKernel(new FuncModule());

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        // ReSharper disable InconsistentNaming
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _kernel.Load(new RubberduckModule(_kernel, (VBE)Application, (AddIn)AddInInst));
                _kernel.Load(new CommandBarsModule(_kernel));
                
                Debug.Print("in OnConnection, ready.");
                var app = _kernel.Get<App>();
                app.Startup();
            }
            catch (Exception exception)
            {
                System.Windows.Forms.MessageBox.Show(exception.ToString(), RubberduckUI.RubberduckLoadFailure, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnStartupComplete(ref Array custom)
        {

        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _kernel.Dispose();
        }
    }
}
