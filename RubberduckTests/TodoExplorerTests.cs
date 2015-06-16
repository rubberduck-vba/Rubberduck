using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;
using Rubberduck.UI;
using Rubberduck.UI.ToDoItems;
using RubberduckTests.Mocks;
using RdMockFacotry = RubberduckTests.Mocks.MockFactory;

namespace RubberduckTests
{
    [TestClass]
    public class TodoExplorerTests
    {
        private Mock<AddIn> _addin;
        private Mock<IToDoExplorerWindow> _view;
        private Mock<Window> _window;
        private MockWindowsCollection _windows;
        private Mock<VBE> _vbe;
        private ConfigurationLoader _loader;
        private ToDoMarker[] _markers;
        private GridViewSort<ToDoItem> _gridViewSorter;

        [TestInitialize]
        public void Intialize()
        {
            _addin = new Mock<AddIn>();
            _view = new Mock<IToDoExplorerWindow>();

            _window = RdMockFacotry.CreateWindowMock();
            _windows = new MockWindowsCollection(_window.Object);

            _loader = new ConfigurationLoader();
            _markers = _loader.GetDefaultTodoMarkers();

            _gridViewSorter = new GridViewSort<ToDoItem>("Priority", false);
        }

        [TestMethod]
        public void TodoPresenter_RefreshUpdatesViewItems()
        {
            var code = @"
Public Sub Bazzer()
    'Todo: Fix the foobarred bazzer.
End Sub";

            var codeModule = RdMockFacotry.CreateCodeModuleMock(code);

            var component = RdMockFacotry.CreateComponentMock("Module1", codeModule.Object, vbext_ComponentType.vbext_ct_StdModule);

            var project = RdMockFacotry.CreateProjectMock("VBAProject", vbext_ProjectProtection.vbext_pp_none);
            
            var componentList = new List<VBComponent>() { component.Object };
            var components = RdMockFacotry.CreateComponentsMock(componentList, project.Object);

            component.SetupGet(c => c.Collection).Returns(components.Object);

            var projectList = new List<VBProject>() {project.Object};

            var projects = RdMockFacotry.CreateProjectsMock(projectList);
            project.SetupGet(p => p.VBComponents).Returns(components.Object);

            _vbe = RdMockFacotry.CreateVbeMock(_windows, projects.Object);

            _view.SetupProperty(v => v.TodoItems);

            var parser = new RubberduckParser();

            var presenter = new ToDoExplorerDockablePresenter(parser, _markers, _vbe.Object, _addin.Object, _view.Object, _gridViewSorter);
          
            //act
            presenter.Refresh();

            //assert
            Assert.AreEqual("Todo: Fix the foobarred bazzer.", _view.Object.TodoItems.First().Description);
        }
    }
}
