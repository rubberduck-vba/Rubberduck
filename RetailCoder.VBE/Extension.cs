using System;
using System.Linq;
using Extensibility;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Windows.Forms;

namespace RetailCoderVBE
{
    [ComVisible(true)]
    [Guid("8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66")]
    [ProgId("RetailCoderVBE.Extension")]
    public class Extension : IDTExtensibility2, IDisposable
    {
        private App _app;

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _app = new App((VBE)Application, (AddIn)AddInInst);
                
                _app.CreateExtUI(); //todo: comment out for release
            }
            catch(Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        public void OnStartupComplete(ref Array custom)
        {
            //_app.CreateExtUI(); //todo: comment out for debugging
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Dispose();
        }

        public void Dispose()
        {
            _app.Dispose();
        }
    }
}
