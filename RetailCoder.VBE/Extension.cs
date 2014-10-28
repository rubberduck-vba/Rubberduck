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
    [ComVisible(true)]
    [Guid("8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66")]
    [ProgId("RetailCoderVBE.Extension")]
    public class Extension : IDTExtensibility2, IDisposable
    {
        private VBE _vbe;
        private TestMenu _testMenu;
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
            Dispose();
        }

        public void Dispose()
        {
            _testSession.Dispose();
        }
    }
}
