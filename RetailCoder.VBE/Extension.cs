using System;
using System.Linq;
using Extensibility;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Collections.Generic;

using RetailCoderVBE.VBIDE;
using RetailCoderVBE.UnitTesting;
using RetailCoderVBE.UnitTesting.UI;

namespace RetailCoderVBE
{
    /* Windows Registry keys
     *
     * [HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64\RetailCoderVBE]
     *  ~> [CommandLineSafe] (DWORD:00000000)
     *  ~> [Description] ("RetailCoderVBE add-in for VBA IDE.")
     *  ~> [LoadBehavior] (DWORD:00000003)
     *  ~> [FriendlyName] ("RetailCoderVBE")
     *
     * [HKEY_CLASSES_ROOT\CLSID\{8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66}]
     *  ~> [@] ("RetailCoderVBE.Extension")
     *  
     * [HKEY_CLASSES_ROOT\CLSID\{8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66}\InprocServer32]
     *  ~> [@] ("mscoree.dll")
     *  ~> [ThreadingModel] ("Both")
     *  ~> [Class] ("RetailCoderVBE.Extension")
     *  ~> [Assembly] ("RetailCoderVBE")
     *  ~> [RuntimeVersion] ("v2.0.50727")
     *  ~> [CodeBase] ("file:///C:\Dev\RetailCoder\RetailCoder.VBE\RetailCoder.VBE\bin\Debug\RetailCoderVBE.dll")
     *
     * [HKEY_CLASSES_ROOT\CLSID\{8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66}\InprocServer32]
     *  ~> [@] ("RetailCoderVBE.Extension")
     *
    */

    [ComVisible(true)]
    [Guid("8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66")]
    [ProgId("RetailCoderVBE.Extension")]
    public class Extension : IDTExtensibility2
    {
        private VBE _vbe;
        private TestMenu _testMenu;
        private TestToolbar _testToolbar;
        private TestSession _testSession;

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
                _vbe = (VBE)Application;
                _testSession = new TestSession(_vbe);
            }
            catch(Exception exception)
            {
                System.Windows.Forms.MessageBox.Show(exception.Message);
            }
        }

        public void OnStartupComplete(ref Array custom)
        {
            CreateTestMenu();
            CreateTestToolbar();
        }

        private void CreateTestMenu()
        {
            _testMenu = new TestMenu();
            _testMenu.OnNewTestClass += OnNewUnitTestModule;
            _testMenu.OnRunAllTests += OnRunAllTests;
            _testMenu.OnRepeatLastRun += OnRepeatLastRun;
            _testMenu.OnRunFailedTests += OnRunFailedTests;
            _testMenu.OnRunPassedTests += OnRunPassedTests;
            _testMenu.OnRunNotRunTests += OnRunNotRunTests;
            _testMenu.OnTestExporer += OnShowTestExplorer;

            _testMenu.Initialize(_vbe);
        }

        void OnShowTestExplorer(object sender, EventArgs e)
        {
            _testSession.ShowExplorer();
        }

        private void CreateTestToolbar()
        {
            _testToolbar = new TestToolbar();
            _testToolbar.OnRunAllTests += OnRunAllTests;
            _testToolbar.OnTestExporer += OnShowTestExplorer;

            _testToolbar.Initialize(_vbe);
        }

        void OnNewExpectedErrorTestMethod(object sender, EventArgs e)
        {
            NewTestMethodCommand.NewExpectedErrorTestMethod(_vbe);
            _testSession.SynchronizeTests();
        }

        void OnNewTestMethod(object sender, EventArgs e)
        {
            NewTestMethodCommand.NewTestMethod(_vbe);
            _testSession.SynchronizeTests();
        }

        void OnRunAllTests(object sender, EventArgs e)
        {
            _testSession.SynchronizeTests();
            _testSession.Run();
        }

        void OnRepeatLastRun(object sender, EventArgs e)
        {
            _testSession.ReRun();
        }

        void OnRunFailedTests(object sender, EventArgs e)
        {
            _testSession.RunFailedTests();
        }

        void OnRunPassedTests(object sender, EventArgs e)
        {
            _testSession.RunPassedTests();
        }

        void OnSynchronizeTests(object sender, EventArgs e)
        {
            _testSession.SynchronizeTests();
        }

        void OnRunNotRunTests(object sender, EventArgs e)
        {
            _testSession.SynchronizeTests();
            _testSession.RunNotRunTests();
        }

        void OnNewUnitTestModule(object sender, EventArgs e)
        {
            NewUnitTestModuleCommand.NewUnitTestModule(_vbe);
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _testSession.Dispose();
        }
    }
}
