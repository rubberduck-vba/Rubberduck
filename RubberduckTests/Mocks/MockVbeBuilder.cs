using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Moq;
using Rubberduck.VBEditor;

namespace RubberduckTests.Mocks
{
    /// <summary>
    /// Builds a mock <see cref="VBE"/>.
    /// </summary>
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

        /// <summary>
        /// Adds a project to the mock VBE.
        /// Use a <see cref="MockProjectBuilder"/> to build the <see cref="project"/>.
        /// </summary>
        /// <param name="project">A mock <see cref="VBProject"/>.</param>
        /// <returns>Returns the <see cref="MockVbeBuilder"/> instance.</returns>
        public MockVbeBuilder AddProject(Mock<VBProject> project)
        {
            project.SetupGet(m => m.VBE).Returns(_vbe.Object);

            _projects.Add(project.Object);

            foreach (var component in _projects.SelectMany(vbProject => vbProject.VBComponents.Cast<VBComponent>()))
            {
                _codePanes.Add(component.CodeModule.CodePane);
            }

            _vbe.SetupGet(vbe => vbe.ActiveVBProject).Returns(project.Object);
            _vbe.SetupGet(vbe => vbe.Version).Returns("7.1");

            _vbProjects = CreateProjectsMock();
            _vbe.SetupGet(m => m.VBProjects).Returns(() => _vbProjects.Object);

            return this;
        }

        /// <summary>
        /// Creates a <see cref="MockProjectBuilder"/> to build a new project.
        /// </summary>
        /// <param name="name">The name of the project to build.</param>
        /// <param name="protection">A value that indicates whether the project is protected.</param>
        public MockProjectBuilder ProjectBuilder(string name, vbext_ProjectProtection protection)
        {
            return ProjectBuilder(name, string.Empty, protection);
        }

        public MockProjectBuilder ProjectBuilder(string name, string filename, vbext_ProjectProtection protection)
        {
            var result = new MockProjectBuilder(name, filename, protection, () => _vbe.Object, this);
            return result;
        }

        /// <summary>
        /// Gets the mock <see cref="VBE"/> instance.
        /// </summary>
        public Mock<VBE> Build()
        {
            return _vbe;
        }

        /// <summary>
        /// Gets a mock <see cref="VBE"/> instance, 
        /// containing a single "TestProject1" <see cref="VBProject"/>
        /// and a single "TestModule1" <see cref="VBComponent"/>, with the specified <see cref="content"/>.
        /// </summary>
        /// <param name="content">The VBA code associated to the component.</param>
        /// <param name="component">The created <see cref="VBComponent"/></param>
        /// <param name="selection"></param>
        /// <returns></returns>
        public Mock<VBE> BuildFromSingleStandardModule(string content, out VBComponent component, Selection selection = new Selection())
        {
            return BuildFromSingleModule(content, vbext_ComponentType.vbext_ct_StdModule, out component, selection);
        }

        public Mock<VBE> BuildFromSingleModule(string content, vbext_ComponentType type, out VBComponent component, Selection selection)
        {
            var builder = ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none);
            builder.AddComponent("TestModule1", type, content, selection);
            var project = builder.Build();
            component = project.Object.VBComponents.Item(0);
            return AddProject(project).Build();
        }

        private Mock<VBE> CreateVbeMock()
        {
            var vbe = new Mock<VBE>();
            var windows = new MockWindowsCollection {VBE = vbe.Object};
            vbe.Setup(m => m.Windows).Returns(() => windows);
            vbe.SetupProperty(m => m.ActiveCodePane);
            vbe.SetupProperty(m => m.ActiveVBProject);
            
            vbe.SetupGet(m => m.SelectedVBComponent).Returns(() => vbe.Object.ActiveCodePane.CodeModule.Parent);
            vbe.SetupGet(m => m.ActiveWindow).Returns(() => vbe.Object.ActiveCodePane.Window);

            var mainWindow = new Mock<Window>();
            mainWindow.Setup(m => m.HWnd).Returns(0);

            vbe.SetupGet(m => m.MainWindow).Returns(() => mainWindow.Object);

            _vbProjects = CreateProjectsMock();
            vbe.SetupGet(m => m.VBProjects).Returns(() => _vbProjects.Object);

            _vbCodePanes = CreateCodePanesMock();
            vbe.SetupGet(m => m.CodePanes).Returns(() => _vbCodePanes.Object);

            return vbe;
        }

        private Mock<VBProjects> CreateProjectsMock()
        {
            var result = new Mock<VBProjects>();

            result.Setup(m => m.GetEnumerator()).Returns(() => _projects.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _projects.GetEnumerator());
            
            result.Setup(m => m.Item(It.IsAny<int>())).Returns<int>(value => _projects.ElementAt(value));
            result.SetupGet(m => m.Count).Returns(() => _projects.Count);


            return result;
        }

        private Mock<CodePanes> CreateCodePanesMock()
        {
            var result = new Mock<CodePanes>();

            result.Setup(m => m.GetEnumerator()).Returns(() => _codePanes.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _codePanes.GetEnumerator());
            
            result.Setup(m => m.Item(It.IsAny<int>())).Returns<int>(value => _codePanes.ElementAt(value));
            result.SetupGet(m => m.Count).Returns(() => _codePanes.Count);

            return result;
        }
    }
}
