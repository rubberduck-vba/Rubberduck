using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Moq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    /// <summary>
    /// Builds a mock <see cref="VBE"/>.
    /// </summary>
    public class MockVbeBuilder
    {
        public const string TestProjectName = "TestProject1";
        public const string TestModuleName = "TestModule1";
        private readonly Mock<IVBE> _vbe;

        #region standard library paths (referenced in all VBA projects hosted in Microsoft Excel)
        public static readonly string LibraryPathVBA = @"C:\PROGRA~2\COMMON~1\MICROS~1\VBA\VBA7.1\VBE7.DLL";      // standard library, priority locked
        public static readonly string LibraryPathMsExcel = @"C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE";   // mock host application, priority locked
        public static readonly string LibraryPathMsOffice = @"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE15\MSO.DLL";
        public static readonly string LibraryPathStdOle = @"C:\Windows\SysWOW64\stdole2.tlb";
        public static readonly string LibraryPathMsForms = @"C:\Windows\SysWOW64\FM20.DLL"; // standard in projects with a UserForm module
        #endregion

        public static readonly string LibraryPathVBIDE = @"C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB";
        public static readonly string LibraryPathScripting = @"C:\Windows\SysWOW64\scrrun.dll";
        public static readonly string LibraryPathRegex = @"C:\Windows\SysWOW64\vbscript.dll\3";
        public static readonly string LibraryPathMsXml = @"C:\Windows\System32\msxml6.dll";
        public static readonly string LibraryPathShDoc = @"C:\Windows\SysWOW64\ieframe.dll";
        public static readonly string LibraryPathAdoDb = @"C:\Program Files (x86)\Common Files\System\ado\msado15.dll";
        public static readonly string LibraryPathAdoRecordset = @"C:\Program Files (x86)\Common Files\System\ado\msador15.dll";

        //private Mock<IWindows> _vbWindows;
        private readonly Windows _windows = new Windows();

        private Mock<IVBProjects> _vbProjects;
        private readonly ICollection<IVBProject> _projects = new List<IVBProject>();
        
        private Mock<ICodePanes> _vbCodePanes;
        private readonly ICollection<ICodePane> _codePanes = new List<ICodePane>(); 

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
        public MockVbeBuilder AddProject(Mock<IVBProject> project)
        {
            project.SetupGet(m => m.VBE).Returns(_vbe.Object);

            _projects.Add(project.Object);

            foreach (var component in _projects.SelectMany(vbProject => vbProject.VBComponents))
            {
                _codePanes.Add(component.CodeModule.CodePane);
            }

            _vbe.SetupGet(vbe => vbe.ActiveVBProject).Returns(project.Object);
            _vbe.SetupGet(vbe => vbe.Version).Returns("7.1");
            _vbe.SetupGet(m => m.VBProjects).Returns(() => _vbProjects.Object);

            return this;
        }

        /// <summary>
        /// Creates a <see cref="MockProjectBuilder"/> to build a new project.
        /// </summary>
        /// <param name="name">The name of the project to build.</param>
        /// <param name="protection">A value that indicates whether the project is protected.</param>
        public MockProjectBuilder ProjectBuilder(string name, ProjectProtection protection)
        {
            return ProjectBuilder(name, string.Empty, protection);
        }

        public MockProjectBuilder ProjectBuilder(string name, string filename, ProjectProtection protection)
        {
            return new MockProjectBuilder(name, filename, protection, () => _vbe.Object, this);
        }

        public MockProjectBuilder ProjectBuilder(string name, string filename, string projectId, ProjectProtection protection)
        {
            return new MockProjectBuilder(name, filename, projectId, protection, () => _vbe.Object, this);
        }

        /// <summary>
        /// Gets the mock <see cref="VBE"/> instance.
        /// </summary>
        public Mock<IVBE> Build()
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
        /// <param name="module">The created <see cref="CodeModule"/></param>
        /// <param name="selection"></param>
        /// <returns></returns>
        public static Mock<IVBE> BuildFromSingleStandardModule(string content, out IVBComponent component, Selection selection = default(Selection), bool referenceStdLibs = false)
        {
            return BuildFromSingleModule(content, TestModuleName, ComponentType.StandardModule, out component, selection, referenceStdLibs);
        }

        public static Mock<IVBE> BuildFromSingleStandardModule(string content, string name, out IVBComponent component, Selection selection = default(Selection), bool referenceStdLibs = false)
        {
            return BuildFromSingleModule(content, name, ComponentType.StandardModule, out component, selection, referenceStdLibs);
        }

        public static Mock<IVBE> BuildFromSingleModule(string content, ComponentType type, out IVBComponent component, Selection selection = default(Selection), bool referenceStdLibs = false)
        {
            return BuildFromSingleModule(content, TestModuleName, type, out component, selection, referenceStdLibs);
        }

        public static Mock<IVBE> BuildFromSingleModule(string content, string name, ComponentType type, out IVBComponent component, Selection selection = default(Selection), bool referenceStdLibs = false)
        {
            var vbeBuilder = new MockVbeBuilder();

            var builder = vbeBuilder.ProjectBuilder(TestProjectName, ProjectProtection.Unprotected);
            builder.AddComponent(name, type, content, selection);

            if (referenceStdLibs)
            {
                builder.AddReference("VBA", LibraryPathVBA, 4, 1, true);
            }

            var project = builder.Build();
            var vbe = vbeBuilder.AddProject(project).Build();

            component = project.Object.VBComponents[0];

            vbe.Object.ActiveVBProject = project.Object;
            vbe.Object.ActiveCodePane = component.CodeModule.CodePane;

            return vbe;
        }

        private Mock<IVBE> CreateVbeMock()
        {
            var vbe = new Mock<IVBE>();
            _windows.VBE = vbe.Object;
            vbe.Setup(m => m.Windows).Returns(() => _windows);
            vbe.SetupProperty(m => m.ActiveCodePane);
            vbe.SetupProperty(m => m.ActiveVBProject);
            
            vbe.SetupGet(m => m.SelectedVBComponent).Returns(() => vbe.Object.ActiveCodePane.CodeModule.Parent);
            vbe.SetupGet(m => m.ActiveWindow).Returns(() => vbe.Object.ActiveCodePane.Window);

            var mainWindow = new Mock<IWindow>();
            mainWindow.Setup(m => m.HWnd).Returns(0);

            vbe.SetupGet(m => m.MainWindow).Returns(() => mainWindow.Object);

            _vbProjects = CreateProjectsMock();
            vbe.SetupGet(m => m.VBProjects).Returns(() => _vbProjects.Object);

            _vbCodePanes = CreateCodePanesMock();
            vbe.SetupGet(m => m.CodePanes).Returns(() => _vbCodePanes.Object);
            
            return vbe;
        }

        private Mock<IVBProjects> CreateProjectsMock()
        {
            var result = new Mock<IVBProjects>();

            result.Setup(m => m.GetEnumerator()).Returns(() => _projects.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _projects.GetEnumerator());
            
            result.Setup(m => m[It.IsAny<int>()]).Returns<int>(value => _projects.ElementAt(value));
            result.SetupGet(m => m.Count).Returns(() => _projects.Count);


            return result;
        }

        private Mock<ICodePanes> CreateCodePanesMock()
        {
            var result = new Mock<ICodePanes>();

            result.Setup(m => m.GetEnumerator()).Returns(() => _codePanes.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _codePanes.GetEnumerator());
            
            result.Setup(m => m[It.IsAny<int>()]).Returns<int>(value => _codePanes.ElementAt(value));
            result.SetupGet(m => m.Count).Returns(() => _codePanes.Count);

            return result;
        }
    }
}
