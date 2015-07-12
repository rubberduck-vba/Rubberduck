using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Moq;

namespace RubberduckTests.Mocks
{
    public class MockVbeBuilder
    {
        private readonly Mock<VBE> _vbe;

        private Mock<VBProjects> _vbProjects;
        private readonly ICollection<VBProject> _projects = new List<VBProject>();
        
        private Mock<CodePanes> _vbCodePanes;
        private readonly ICollection<CodePane> _codePanes = new List<CodePane>(); 

        public MockVbeBuilder()
        {
            _vbe = CreateVbeMock();
        }

        public MockVbeBuilder AddProject(Mock<VBProject> project)
        {
            project.SetupGet(m => m.VBE).Returns(_vbe.Object);
            _projects.Add(project.Object);

            foreach (var component in _projects.SelectMany(vbProject => vbProject.VBComponents.Cast<VBComponent>()))
            {
                _codePanes.Add(component.CodeModule.CodePane);
            }

            return this;
        }

        public MockProjectBuilder ProjectBuilder(string name, vbext_ProjectProtection protection)
        {
            var result = new MockProjectBuilder(name, protection, () => _vbe.Object);
            return result;
        }

        public Mock<VBE> Build()
        {
            return _vbe;
        }

        private Mock<VBE> CreateVbeMock()
        {
            var vbe = new Mock<VBE>();
            var windows = new MockWindowsCollection {VBE = vbe.Object};
            vbe.Setup(m => m.Windows).Returns(windows);
            vbe.SetupProperty(m => m.ActiveCodePane);
            vbe.SetupProperty(m => m.ActiveVBProject);
            
            vbe.SetupGet(m => m.SelectedVBComponent).Returns(() => vbe.Object.ActiveCodePane.CodeModule.Parent);
            vbe.SetupGet(m => m.ActiveWindow).Returns(() => vbe.Object.ActiveCodePane.Window);

            var mainWindow = new Mock<Window>();
            mainWindow.Setup(m => m.HWnd).Returns(0);

            vbe.SetupGet(m => m.MainWindow).Returns(mainWindow.Object);

            _vbProjects = CreateProjectsMock();
            _vbe.SetupGet(m => m.VBProjects).Returns(() => _vbProjects.Object);

            _vbCodePanes = CreateCodePanesMock();
            _vbe.SetupGet(m => m.CodePanes).Returns(() => _vbCodePanes.Object);

            return vbe;
        }

        private Mock<VBProjects> CreateProjectsMock()
        {
            var result = new Mock<VBProjects>();

            result.Setup(m => m.GetEnumerator()).Returns(_projects.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(_projects.GetEnumerator());
            
            result.Setup(m => m.Item(It.IsAny<int>())).Returns<int>(value => _projects.ElementAt(value));
            result.SetupGet(m => m.Count).Returns(_projects.Count);


            return result;
        }

        private Mock<CodePanes> CreateCodePanesMock()
        {
            var result = new Mock<CodePanes>();

            result.Setup(m => m.GetEnumerator()).Returns(_codePanes.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(_codePanes.GetEnumerator());
            
            result.Setup(m => m.Item(It.IsAny<int>())).Returns<int>(value => _codePanes.ElementAt(value));
            result.SetupGet(m => m.Count).Returns(_codePanes.Count);

            return result;
        }
    }
}
