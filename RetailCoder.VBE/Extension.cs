using Extensibility;
using Ninject;
using Ninject.Extensions.Factory;
using Rubberduck.Common.WinAPI;
using Rubberduck.Root;
using Rubberduck.UI;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Threading;
using Ninject.Extensions.Interception;
using NLog;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck
{
    /// <remarks>
    /// Special thanks to Carlos Quintero (MZ-Tools) for providing the general structure here.
    /// </remarks>
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgId)]
    [EditorBrowsable(EditorBrowsableState.Never)]
    // ReSharper disable once InconsistentNaming // note: underscore prefix hides class from COM API
    public class _Extension : IDTExtensibility2
    {
        private const string ClassId = "8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66";
        private const string ProgId = "Rubberduck.Extension";

        private IVBE _ide;
        private IAddIn _addin;
        private bool _isInitialized;
        private bool _isBeginShutdownExecuted;

        private IKernel _kernel;
        private App _app;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public void OnAddInsUpdate(ref Array custom) { }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                if (Application is Microsoft.Vbe.Interop.VBE)
                {
                    var vbe = (Microsoft.Vbe.Interop.VBE) Application;                  
                    _ide = new VBEditor.SafeComWrappers.VBA.VBE(vbe);
                    VBENativeServices.HookEvents(_ide);
                    
                    var addin = (Microsoft.Vbe.Interop.AddIn)AddInInst;
                    _addin = new VBEditor.SafeComWrappers.VBA.AddIn(addin) { Object = this };
                }
                else if (Application is Microsoft.VB6.Interop.VBIDE.VBE)
                {
                    var vbe = Application as Microsoft.VB6.Interop.VBIDE.VBE;
                    _ide = new VBEditor.SafeComWrappers.VB6.VBE(vbe);

                    var addin = (Microsoft.VB6.Interop.VBIDE.AddIn) AddInInst;
                    _addin = new VBEditor.SafeComWrappers.VB6.AddIn(addin);
                }


                switch (ConnectMode)
                {
                    case ext_ConnectMode.ext_cm_Startup:
                        // normal execution path - don't initialize just yet, wait for OnStartupComplete to be called by the host.
                        break;
                    case ext_ConnectMode.ext_cm_AfterStartup:
                        InitializeAddIn();
                        break;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        Assembly LoadFromSameFolder(object sender, ResolveEventArgs args)
        {
            var folderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;
            var assemblyPath = Path.Combine(folderPath, new AssemblyName(args.Name).Name + ".dll");
            if (!File.Exists(assemblyPath))
            {
                return null;
            }

            var assembly = Assembly.LoadFile(assemblyPath);
            return assembly;
        }

        public void OnStartupComplete(ref Array custom)
        {
            InitializeAddIn();
        }

        public void OnBeginShutdown(ref Array custom)
        {
            _isBeginShutdownExecuted = true;
            ShutdownAddIn();
        }

        // ReSharper disable InconsistentNaming
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {            
            switch (RemoveMode)
            {
                case ext_DisconnectMode.ext_dm_UserClosed:
                    ShutdownAddIn();
                    break;

                case ext_DisconnectMode.ext_dm_HostShutdown:
                    if (_isBeginShutdownExecuted)
                    {
                        // this is the normal case: nothing to do here, we already ran ShutdownAddIn.
                    }
                    else
                    {
                        // some hosts do not call OnBeginShutdown: this mitigates it.
                        ShutdownAddIn();
                    }
                    break;
            }
        }

        private void InitializeAddIn()
        {
            if (_isInitialized)
            {
                // The add-in is already initialized. See:
                // The strange case of the add-in initialized twice
                // http://msmvps.com/blogs/carlosq/archive/2013/02/14/the-strange-case-of-the-add-in-initialized-twice.aspx
                return;
            }

            var configLoader = new XmlPersistanceService<GeneralSettings>
            {
                FilePath =
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "Rubberduck", "rubberduck.config")
            };
            var configProvider = new GeneralConfigProvider(configLoader);
            
            var settings = configProvider.Create();
            if (settings != null)
            {
                try
                {
                    var cultureInfo = CultureInfo.GetCultureInfo(settings.Language.Code);
                    Dispatcher.CurrentDispatcher.Thread.CurrentUICulture = cultureInfo;
                }
                catch (CultureNotFoundException)
                {
                }
            }
            else
            {
                Debug.Assert(false, "Settings could not be initialized.");
            }

            Splash splash = null;
            if (settings.ShowSplash)
            {
                splash = new Splash
                {
                    // note: IVersionCheck.CurrentVersion could return this string.
                    Version = string.Format("version {0}", Assembly.GetExecutingAssembly().GetName().Version)
                };
                splash.Show();
                splash.Refresh();
            }

            try
            {
                Startup();
            }
            catch (Win32Exception)
            {
                System.Windows.Forms.MessageBox.Show(RubberduckUI.RubberduckReloadFailure_Message, RubberduckUI.RubberduckReloadFailure_Title,
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception exception)
            {
                _logger.Fatal(exception);
                System.Windows.Forms.MessageBox.Show(exception.ToString(), RubberduckUI.RubberduckLoadFailure,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (splash != null)
                {
                    splash.Dispose();
                }
            }
        }

        private void Startup()
        {
            var currentDomain = AppDomain.CurrentDomain;
            currentDomain.AssemblyResolve += LoadFromSameFolder;

            _kernel = new StandardKernel(new NinjectSettings {LoadExtensions = true}, new FuncModule(), new DynamicProxyModule());
            _kernel.Load(new RubberduckModule(_ide, _addin));

            _app = _kernel.Get<App>();
            _app.Startup();

            _isInitialized = true;
        }

        private void ShutdownAddIn()
        {
            VBENativeServices.UnhookEvents();

            var currentDomain = AppDomain.CurrentDomain;
            currentDomain.AssemblyResolve -= LoadFromSameFolder;

            User32.EnumChildWindows(_ide.MainWindow.Handle(), EnumCallback, new IntPtr(0));

            if (_app != null)
            {
                _app.Shutdown();
                _app = null;
            }

            if (_kernel != null)
            {
                _kernel.Dispose();
                _kernel = null;
            }

            try
            {
                _ide.Release();
            }
            catch (Exception e)
            {
                _logger.Error(e);
            }

            GC.WaitForPendingFinalizers();
            _isInitialized = false;
        }

        private static int EnumCallback(IntPtr hwnd, IntPtr lparam)
        {
            User32.SendMessage(hwnd, WM.RUBBERDUCK_SINKING, IntPtr.Zero, IntPtr.Zero);
            return 1;
        }
    }
}
