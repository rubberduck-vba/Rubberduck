using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Moq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    /// <summary>
    /// Builds a mock <see cref="IVBE"/>.
    /// </summary>
    [SuppressMessage("Microsoft.Design", "CA1001")] //CA1001 is complaining about RubberduckTests.Mocks.Windows, which doesn't need to be disposed in this context.
    public class MockVbeBuilder
    {
        public const string TestProjectName = "TestProject1";
        public const string TestModuleName = "TestModule1";
        private readonly Mock<IVBE> _vbe;
        private readonly Mock<IVbeEvents> _vbeEvents;

        #region standard library paths (referenced in all VBA projects hosted in Microsoft Excel)
        public static readonly string LibraryPathVBA = @"C:\PROGRA~1\COMMON~1\MICROS~1\VBA\VBA7.1\VBE7.DLL";      // standard library, priority locked
        public static readonly string LibraryPathMsExcel = @"C:\Program Files\Microsoft Office\Office15\EXCEL.EXE";   // mock host application, priority locked
        public static readonly string LibraryPathMsOffice = @"C:\Program Files\Common Files\Microsoft Shared\OFFICE15\MSO.DLL";
        public static readonly string LibraryPathStdOle = @"C:\Windows\System32\stdole2.tlb";
        public static readonly string LibraryPathMsForms = @"C:\Windows\system32\FM20.DLL"; // standard in projects with a UserForm module
        #endregion

        public static readonly string LibraryPathVBIDE = @"C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB";
        public static readonly string LibraryPathScripting = @"C:\Windows\System32\scrrun.dll";
        public static readonly string LibraryPathRegex = @"C:\Windows\System32\vbscript.dll\3";
        public static readonly string LibraryPathMsXml = @"C:\Windows\System32\msxml6.dll";
        public static readonly string LibraryPathShDoc = @"C:\Windows\System32\ieframe.dll";
        public static readonly string LibraryPathAdoDb = @"C:\Program Files\Common Files\System\ado\msado15.dll";
        public static readonly string LibraryPathAdoRecordset = @"C:\Program Files\Common Files\System\ado\msador15.dll";

        public static readonly Dictionary<string, string> LibraryPaths = new Dictionary<string, string>
        {
            ["VBA"] = LibraryPathVBA,
            ["Excel"] = LibraryPathMsExcel,
            ["Office"] = LibraryPathMsOffice,
            ["stdole"] = LibraryPathStdOle,
            ["MSForms"] = LibraryPathMsForms,
            ["VBIDE"] = LibraryPathVBIDE,
            ["Scripting"] = LibraryPathScripting,
            ["VBScript_RegExp_55"] = LibraryPathRegex,
            ["MSXML2"] = LibraryPathMsXml,
            ["SHDocVw"] = LibraryPathShDoc,
            ["ADODB"] = LibraryPathAdoDb,
            ["ADOR"] = LibraryPathAdoRecordset
        };

        private static readonly Dictionary<string, Func<MockProjectBuilder, MockProjectBuilder>> AddReference = new Dictionary<string, Func<MockProjectBuilder, MockProjectBuilder>>
        {
            ["Excel"] = (MockProjectBuilder builder) => builder.AddReference("Excel", LibraryPathMsExcel, 1, 8, true),
            ["VBA"] = (MockProjectBuilder builder) => builder.AddReference("VBA", LibraryPathVBA, 4, 2, true),
            ["Scripting"] = (MockProjectBuilder builder) => builder.AddReference("Scripting", LibraryPathScripting, 1, 0, true),
            ["ADODB"] = (MockProjectBuilder builder) => builder.AddReference("ADODB", LibraryPathAdoDb, 6, 1, false),
            ["MSForms"] = (MockProjectBuilder builder) => builder.AddReference("MSForms", LibraryPathMsForms, 2, 0, true),
        };

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
        /// <param name="project">A mock <see cref="IVBProject"/>.</param>
        /// <returns>Returns the <see cref="MockVbeBuilder"/> instance.</returns>
        public MockVbeBuilder AddProject(Mock<IVBProject> project)
        {
            project.SetupGet(m => m.VBE).Returns(_vbe.Object);

            _projects.Add(project.Object);

            var allCodePanes = _projects.SelectMany(vbProject => vbProject.VBComponents)
                .Select(component => component.CodeModule.CodePane);
            AddOpenCodePanes(allCodePanes);

            _vbe.SetupGet(vbe => vbe.ActiveVBProject).Returns(project.Object);
            _vbe.SetupGet(m => m.VBProjects).Returns(() => _vbProjects.Object);

            return this;
        }

        /// <summary>
        /// Creates a <see cref="MockProjectBuilder"/> to build a new project.
        /// </summary>
        /// <param name="name">The name of the project to build.</param>
        /// <param name="protection">A value that indicates whether the project is protected.</param>
        public MockProjectBuilder ProjectBuilder(string name, ProjectProtection protection, ProjectType projectType = ProjectType.HostProject)
        {
            return ProjectBuilder(name, string.Empty, protection, projectType);
        }

        public MockProjectBuilder ProjectBuilder(string name, string filename, ProjectProtection protection, ProjectType projectType = ProjectType.HostProject)
        {
            return new MockProjectBuilder(name, filename, protection, projectType, () => _vbe.Object, this);
        }

        public MockProjectBuilder ProjectBuilder(string name, string filename, string projectId, ProjectProtection protection, ProjectType projectType = ProjectType.HostProject)
        {
            return new MockProjectBuilder(name, filename, projectId, protection, projectType, () => _vbe.Object, this);
        }

        /// <summary>
        /// Gets the mock <see cref="IVBE"/> instance.
        /// </summary>
        public Mock<IVBE> Build()
        {
            _vbe.SetupGet(vbe => vbe.Version).Returns("7.1");
            return _vbe;
        }

        /// <summary>
        /// Gets a mock <see cref="IVBE"/> instance, 
        /// containing a single "TestProject1" <see cref="IVBProject"/>
        /// and a single "TestModule1" <see cref="IVBComponent"/>, with the specified <see cref="content"/>.
        /// </summary>
        /// <param name="content">The VBA code associated to the component.</param>
        /// <param name="component">The created <see cref="IVBComponent"/></param>
        /// <param name="selection">Specifies user selection in the editor.</param>
        /// <param name="referenceStdLibs">Specifies whether standard libraries are referenced.</param>
        /// <returns></returns>
        public static Mock<IVBE> BuildFromSingleStandardModule(string content, out IVBComponent component, Selection selection = default, bool referenceStdLibs = false)
        {
            return BuildFromSingleModule(content, TestModuleName, ComponentType.StandardModule, out component, selection, referenceStdLibs);
        }

        public static Mock<IVBE> BuildFromSingleStandardModule(string content, string name, out IVBComponent component, Selection selection = default, bool referenceStdLibs = false)
        {
            return BuildFromSingleModule(content, name, ComponentType.StandardModule, out component, selection, referenceStdLibs);
        }

        public static Mock<IVBE> BuildFromSingleModule(string content, ComponentType type, out IVBComponent component, Selection selection = default, bool referenceStdLibs = false)
        {
            return BuildFromSingleModule(content, TestModuleName, type, out component, selection, referenceStdLibs);
        }

        public static Mock<IVBE> BuildFromSingleModule(string content, string name, ComponentType type, out IVBComponent component, Selection selection = default, bool referenceStdLibs = false)
        {
            var vbeBuilder = new MockVbeBuilder();

            var builder = vbeBuilder.ProjectBuilder(TestProjectName, ProjectProtection.Unprotected);
            builder.AddComponent(name, type, content, selection);

            if (referenceStdLibs)
            {
                builder.AddReference("VBA", LibraryPathVBA, 4, 2, true);
            }

            var project = builder.Build();
            var vbe = vbeBuilder.AddProject(project).Build();

            component = project.Object.VBComponents[0];

            vbe.Object.ActiveVBProject = project.Object;
            vbe.Object.ActiveCodePane = component.CodeModule.CodePane;

            return vbe;
        }

        /// <summary>
        /// Builds a mock VBE containing a single "TestProject1" and multiple standard modules.
        /// </summary>
        public static Mock<IVBE> BuildFromStdModules(params (string name, string content)[] modules)
        {
            return BuildFromModules(modules.Select(tpl => (tpl.name, tpl.content, ComponentType.StandardModule)));
        }

        /// <summary>
        /// Builds a mock VBE containing a single "TestProject1" with multiple modules.
        /// </summary>
        public static Mock<IVBE> BuildFromModules(params (string name, string content, ComponentType componentType)[] modules)
        {
            return BuildFromModules((IEnumerable<(string name, string content, ComponentType componentType)>)modules);
        }

        /// <summary>
        /// Builds a mock VBE containing a single "TestProject1" with multiple modules.
        /// </summary>
        public static Mock<IVBE> BuildFromModules(IEnumerable<(string name, string content, ComponentType componentType)> modules)
            => BuildFromModules(modules, Enumerable.Empty<string>());

        /// <summary>
        /// Builds a mock VBE containing a single "TestProject1" with one module and one or more libraries.
        /// </summary>
        public static Mock<IVBE> BuildFromModules((string name, string content, ComponentType componentType) module, params string[] libraries)
            => BuildFromModules(new (string, string, ComponentType)[] { module }, libraries);

        /// <summary>
        /// Builds a mock VBE containing a single "TestProject1" with multiple modules and libraries.
        /// </summary>
        public static Mock<IVBE> BuildFromModules(IEnumerable<(string name, string content, ComponentType componentType)> modules, IEnumerable<string> libraryNames)
            => BuildFromModules(TestProjectName, modules, libraryNames);

        /// <summary>
        /// Builds a mock VBE containing one project with multiple modules and libraries.
        /// </summary>
        public static Mock<IVBE> BuildFromModules(string projectName, IEnumerable<(string name, string content, ComponentType componentType)> modules, IEnumerable<string> libraryNames)
        {
            var vbeBuilder = new MockVbeBuilder();

            var builder = vbeBuilder.ProjectBuilder(projectName, ProjectProtection.Unprotected);
            foreach (var (name, content, componentType) in modules)
            {
                builder.AddComponent(name, componentType, content);
            }

            foreach (var name in libraryNames)
            {
                AddReference[name](builder);
            }

            var project = builder.Build();
            var vbe = vbeBuilder.AddProject(project).Build();

            var component = project.Object.VBComponents[0];

            vbe.Object.ActiveVBProject = project.Object;
            vbe.Object.ActiveCodePane = component.CodeModule.CodePane;

            return vbe;
        }

        public MockVbeBuilder SetOpenCodePanes(IEnumerable<ICodePane> openCodePanes)
        {
            _codePanes.Clear();
            AddOpenCodePanes(openCodePanes);
            return this;
        }

        public MockVbeBuilder AddOpenCodePanes(IEnumerable<ICodePane> openCodePanes)
        {
            foreach (var codePane in openCodePanes)
            {
                _codePanes.Add(codePane);
            }

            return this;
        }

        private Mock<IVBE> CreateVbeMock()
        {
            var vbe = new Mock<IVBE>();
            _windows.VBE = vbe.Object;
            vbe.Setup(m => m.Dispose());
            vbe.SetupReferenceEqualityIncludingHashCode();
            vbe.Setup(m => m.Windows).Returns(() => _windows);
            vbe.SetupProperty(m => m.ActiveCodePane);
            vbe.SetupProperty(m => m.ActiveVBProject);
            
            vbe.SetupGet(m => m.SelectedVBComponent).Returns(() => vbe.Object.ActiveCodePane?.CodeModule?.Parent);
            vbe.Setup(m => m.GetActiveSelection()).Returns(() => vbe.Object.ActiveCodePane?.GetQualifiedSelection());
            vbe.SetupGet(m => m.ActiveWindow).Returns(() => vbe.Object.ActiveCodePane.Window);

            var mainWindow = new Mock<IWindow>();
            mainWindow.Setup(m => m.HWnd).Returns(0);

            vbe.SetupGet(m => m.MainWindow).Returns(() => mainWindow.Object);

            _vbProjects = CreateProjectsMock();
            vbe.SetupGet(m => m.VBProjects).Returns(() => _vbProjects.Object);

            _vbCodePanes = CreateCodePanesMock();
            vbe.SetupGet(m => m.CodePanes).Returns(() => _vbCodePanes.Object);

            var commandBars = DummyCommandBars();
            vbe.SetupGet(m => m.CommandBars).Returns(() => commandBars);

            vbe.Setup(m => m.IsInDesignMode).Returns(true);

            return vbe;
        }

        private static ICommandBars DummyCommandBars()
        {
            var commandBars = new Mock<ICommandBars>(); 
 
            var dummyCommandBar = DummyCommandBar();

            commandBars.SetupGet(m => m[It.IsAny<int>()]).Returns(dummyCommandBar);

            return commandBars.Object;
        }

        private static ICommandBar DummyCommandBar()
        {
            var commandBar = new Mock<ICommandBar>();

            var commandBarControlCollection = new List<ICommandBarControl>();
            var commandBarControls = CommandBarControlsFromCollection(commandBarControlCollection);

            commandBar.SetupGet(m => m.Controls).Returns(commandBarControls.Object);

            return commandBar.Object;
        }

        private static Mock<ICommandBarControls> CommandBarControlsFromCollection(IList<ICommandBarControl> commandBarControlCollection)
        {
            var commandBarControls = new Mock<ICommandBarControls>();

            commandBarControls.Setup(m => m.GetEnumerator()).Returns(() => commandBarControlCollection.GetEnumerator());
            commandBarControls.As<IEnumerable>().Setup(m => m.GetEnumerator())
                .Returns(() => commandBarControlCollection.GetEnumerator());

            commandBarControls.Setup(m => m[It.IsAny<int>()])
                .Returns<int>(value => commandBarControlCollection.ElementAt(value));
            commandBarControls.SetupGet(m => m.Count).Returns(() => commandBarControlCollection.Count);
            return commandBarControls;
        }

        private Mock<IVBProjects> CreateProjectsMock()
        {
            var result = new Mock<IVBProjects>();

            result.Setup(m => m.Dispose());
            result.SetupReferenceEqualityIncludingHashCode();

            result.Setup(m => m.GetEnumerator()).Returns(() => _projects.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _projects.GetEnumerator());
            
            result.Setup(m => m[It.IsAny<int>()]).Returns<int>(value => _projects.ElementAt(value));
            result.SetupGet(m => m.Count).Returns(() => _projects.Count);

            result.Setup(m => m.Add(It.IsAny<ProjectType>()))
                .Returns((ProjectType pt) =>
                {
                    var projectBuilder = ProjectBuilder("test", ProjectProtection.Unprotected);
                    var project = projectBuilder.Build();
                    project.Object.AssignProjectId();
                    AddProject(project);
                    return project.Object;
                });
            result.Setup(m => m.Open(It.IsAny<string>()))
                .Returns((string path) =>
                {
                    var projectBuilder = ProjectBuilder("openedTestProject", path, ProjectProtection.Locked);
                    var projectId = QualifiedModuleName.GetProjectId("openedTestProject", path);
                    var project = projectBuilder.Build();
                    project.Setup(m => m.ProjectId).Returns(projectId);
                    AddProject(project);
                    return project.Object;
                });
            result.Setup(m => m.Remove(It.IsAny<IVBProject>())).Callback((IVBProject proj) => _projects.Remove(proj));

            return result;
        }

        private Mock<ICodePanes> CreateCodePanesMock()
        {
            var result = new Mock<ICodePanes>();

            result.Setup(m => m.Dispose());
            result.SetupReferenceEqualityIncludingHashCode();

            result.Setup(m => m.GetEnumerator()).Returns(() => _codePanes.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _codePanes.GetEnumerator());
            
            result.Setup(m => m[It.IsAny<int>()]).Returns<int>(value => _codePanes.ElementAt(value));
            result.SetupGet(m => m.Count).Returns(() => _codePanes.Count);

            return result;
        }

        public Mock<IVBProjects> MockProjectsCollection => _vbProjects;
    }
}
