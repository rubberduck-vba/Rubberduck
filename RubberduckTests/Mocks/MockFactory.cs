using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Moq;

namespace RubberduckTests.Mocks
{
    static class MockFactory
    {
        /// <summary>
        /// Creates a mock <see cref="Window"/> that is particularly useful for passing into <see cref="MockWindowsCollection"/>'s ctor.
        /// </summary>
        /// <returns>
        /// A <see cref="Mock{Window}"/>that has all the properties needed for <see cref="Rubberduck.UI.DockablePresenterBase"/> pre-setup.
        /// </returns>
        internal static Mock<Window> CreateWindowMock()
        {
            var window = new Mock<Window>();
            window.SetupProperty(w => w.Visible, false);
            window.SetupGet(w => w.LinkedWindows).Returns((LinkedWindows) null);
            window.SetupProperty(w => w.Height);
            window.SetupProperty(w => w.Width);

            return window;
        }

        /// <summary>
        /// Creates a mock <see cref="Window"/> with it's <see cref="Window.Caption"/> propery set up.
        /// </summary>
        /// <param name="caption">The value to return from <see cref="Window.Caption"/>.</param>
        /// <returns>
        /// A <see cref="Mock{Window}"/>that has all the properties needed for <see cref="Rubberduck.UI.DockablePresenterBase"/> pre-setup.
        /// </returns>
        internal static Mock<Window> CreateWindowMock(string caption)
        {
            var window = CreateWindowMock();
            window.SetupGet(w => w.Caption).Returns(caption);

            return window;
        }

        internal static Mock<VBE> CreateVbeMock()
        {
            return CreateVbeMock(new MockWindowsCollection());
        }

        /// <summary>
        /// Creates a new <see cref="Mock{VBE}"/> that returns the <see cref="Windows"/> collection argument out of the Windows property.
        /// </summary>
        /// <param name="windows">
        /// A <see cref="MockWindowsCollection"/> is expected. 
        /// Other objects implementing the<see cref="Windows"/> interface could cause issues.
        /// </param>
        /// <returns></returns>
        internal static Mock<VBE> CreateVbeMock(MockWindowsCollection windows)
        {
            var vbe = new Mock<VBE>();
            windows.VBE = vbe.Object;
            vbe.Setup(v => v.Windows).Returns(windows);

            //setting up a main window lets the native window functions fun
            var mainWindow = new Mock<Window>();
            mainWindow.Setup(w => w.HWnd).Returns(0);

            vbe.SetupGet(v => v.MainWindow).Returns(mainWindow.Object);

            return vbe;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{VBE}"/> with the <see cref="VBE.Windows"/> and <see cref="VBE.VBProjects"/> properties setup.
        /// </summary>
        /// <param name="windows">
        /// A <see cref="MockWindowsCollection"/> is expected. 
        /// Other objects implementing the<see cref="Windows"/> interface could cause issues.
        /// </param>
        /// <param name="projects"><see cref="VBProjects"/> collecction.</param>
        /// <returns></returns>
        internal static Mock<VBE> CreateVbeMock(MockWindowsCollection windows, VBProjects projects)
        {
            var vbe = CreateVbeMock(windows);
            vbe.SetupGet(v => v.VBProjects).Returns(projects);

            return vbe;
        }

        /// <summary>
        /// Creates a "selectable" <see cref="Mock{CodePane}"/>.
        /// </summary>
        /// <param name="vbe">Returned back from the <see cref="CodePane.VBE"/> property.</param>
        /// <param name="name">The caption of the window object that will be created for this code pane.</param>
        /// <returns></returns>
        internal static Mock<CodePane> CreateCodePaneMock(Mock<VBE> vbe, string name)
        {
            var windows = vbe.Object.Windows as MockWindowsCollection;
            if (windows == null)
            {
                return null;
            }

            var codePane = new Mock<CodePane>();
            var window = windows.CreateWindow(name);
            windows.Add(window);

            codePane.Setup(p => p.SetSelection(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<int>(), It.IsAny<int>()));
            codePane.Setup(p => p.Show());
            codePane.SetupGet(p => p.VBE).Returns(vbe.Object);
            codePane.SetupGet(p => p.Window).Returns(window);
            return codePane;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{CodeModule}"/> setup to appropriately mimic getting and modifying code contained in the <see cref="CodeModule"/>.
        /// </summary>
        /// <param name="code">A block of VBA code.</param>
        /// <returns></returns>
        internal static Mock<CodeModule> CreateCodeModuleMock(string code)
        {
            var lines = code.Split(new[] {Environment.NewLine}, StringSplitOptions.None).ToList();

            var codeModule = new Mock<CodeModule>();
            codeModule.SetupGet(c => c.CountOfLines).Returns(lines.Count);

            // ReSharper disable once UseIndexedProperty
            // No R#, the indexed property breaks the expression. I tried that first.
            codeModule.Setup(m => m.get_Lines(It.IsAny<int>(), It.IsAny<int>()))
                .Returns<int, int>((start, count) => String.Join(Environment.NewLine, lines.Skip(start - 1).Take(count)));

            codeModule.Setup(m => m.ReplaceLine(It.IsAny<int>(), It.IsAny<string>()))
                .Callback<int, string>((index, str) => lines[index - 1] = str);

            codeModule.Setup(m => m.DeleteLines(It.IsAny<int>(), It.IsAny<int>()))
                .Callback<int, int>((index, count) => lines.RemoveRange(index - 1, count));

            codeModule.Setup(m => m.InsertLines(It.IsAny<int>(), It.IsAny<string>()))
                .Callback<int, string>((index, newLine) => lines.Insert(index - 1, newLine));
                
            return codeModule;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{CodeModule}"/> setup to appropriately mimic getting and modifying code contained in the <see cref="CodeModule"/>.
        /// </summary>
        /// <param name="code">A block of VBA code.</param>
        /// <param name="codePane">Returned back from the <see cref="CodeModule.CodePane"/> property.</param>
        /// <returns></returns>
        internal static Mock<CodeModule> CreateCodeModuleMock(string code, Mock<CodePane> codePane, Mock<VBE> vbe)
        {
            var codeModule = CreateCodeModuleMock(code);
            codeModule.SetupGet(m => m.CodePane).Returns(codePane.Object);
            codeModule.SetupGet(m => m.VBE).Returns(vbe.Object);

            return codeModule;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{VBComponent}"/>.
        /// </summary>
        /// <param name="name">The name to return from the <see cref="VBComponent.Name"/> property.</param>
        /// <param name="codeModule">The <see cref="CodeModule"/> to return from the CodeModule property.</param>
        /// <param name="componentType">
        /// The type of component to be simulated.
        /// Use vbext_ct_StdModule for standard modules.
        /// Use vbext_ct_ClassModule for classes.
        /// vbext_ct_ActiveXDesigner is invalid for the VBE.
        /// </param>
        /// <returns></returns>
        internal static Mock<VBComponent> CreateComponentMock(string name, CodeModule codeModule, vbext_ComponentType componentType, Mock<VBE> vbe)
        {
            var component = new Mock<VBComponent>();
            component.SetupProperty(m => m.Name, name);
            component.SetupGet(m => m.CodeModule).Returns(codeModule);
            component.SetupGet(m => m.Type).Returns(componentType);
            component.SetupGet(m => m.VBE).Returns(vbe.Object);
            return component;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{VBComponents}"/> that can be iterated over as an <see cref="IEnumerable"/>.
        /// </summary>
        /// <param name="componentList">The collection to be iterated over.</param>
        /// <returns></returns>
        internal static Mock<VBComponents> CreateComponentsMock(IEnumerable<VBComponent> componentList)
        {
            var components = new Mock<VBComponents>();
            components.Setup(c => c.GetEnumerator()).Returns(componentList.GetEnumerator());
            components.As<IEnumerable>().Setup(c => c.GetEnumerator()).Returns(componentList.GetEnumerator());

            return components;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{VBComponents}"/> that can be iterated over as an <see cref="IEnumerable"/>.
        /// </summary>
        /// <param name="componentList">The collection to be iterated over.</param>
        /// <param name="project">The <see cref="VBComponents.Parent"/> property.</param>
        /// <returns></returns>
        internal static Mock<VBComponents> CreateComponentsMock(ICollection<Mock<VBComponent>> componentList, VBProject project)
        {
            var items = componentList.Select(item => item.Object);
            var components = CreateComponentsMock(items);

            foreach (var mock in componentList)
            {
                mock.SetupGet(m => m.Collection).Returns(components.Object);
            }

            return components;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{VBProject}"/>.
        /// </summary>
        /// <param name="name">The <see cref="VBProject.Name"/> property.</param>
        /// <param name="protectionLevel">
        /// The <see cref="VBProject.Protection"/> property.
        /// Use vbext_pp_none to simulate a normal project.
        /// Use vbext_pp_locked to simulate a password protected, or otherwise unavailable, project.
        /// </param>
        /// <returns></returns>
        internal static Mock<VBProject> CreateProjectMock(string name, vbext_ProjectProtection protectionLevel)
        {
            var project = new Mock<VBProject>();
            project.SetupProperty(p => p.Name, name);
            project.SetupGet(p => p.Protection).Returns(protectionLevel);
            return project;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{VBProjects}"/> that can be iterated over as an <see cref="IEnumerable"/>.
        /// </summary>
        /// <param name="projectList">The collection to be iterated over.</param>
        /// <returns></returns>
        internal static Mock<VBProjects> CreateProjectsMock(ICollection<VBProject> projectList)
        {
            var projects = new Mock<VBProjects>();
            projects.Setup(p => p.GetEnumerator()).Returns(projectList.GetEnumerator());
            projects.As<IEnumerable>().Setup(p => p.GetEnumerator()).Returns(projectList.GetEnumerator());

            return projects;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{Reference}"/>.
        /// </summary>
        /// <param name="name">The see<see cref="Reference.Name"/>.</param>
        /// <param name="filePath">The <see cref="Reference.FullPath"/> filepath.</param>
        /// <returns></returns>
        internal static Mock<Reference> CreateMockReference(string name, string filePath)
        {
            var reference = new Mock<Reference>();
            reference.SetupGet(r => r.Name).Returns(name);
            reference.SetupGet(r => r.FullPath).Returns(filePath);

            return reference;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{References}"/> collection that can be iterated over as an <see cref="IEnumerable"/>.
        /// </summary>
        /// <param name="referenceList">The collection to be iterated over.</param>
        /// <returns></returns>
        internal static Mock<References> CreateReferencesMock(List<Reference> referenceList)
        {
            var references = new Mock<References>();
            references.Setup(r => r.GetEnumerator()).Returns(referenceList.GetEnumerator());
            references.As<IEnumerable>().Setup(r => r.GetEnumerator()).Returns(referenceList.GetEnumerator());
            return references;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{Project}"/> that is set up with a <see cref="References"/> collection.
        /// </summary>
        /// <param name="name">The <see cref="VBProject"/> <see cref="VBProject.Name"/>.</param>
        /// <param name="references">The <see cref="References"/> collection.</param>
        /// <returns></returns>
        internal static Mock<VBProject> CreateProjectMock(string name, Mock<References> references)
        {
            var project = new Mock<VBProject>();
            project.SetupProperty(p => p.Name, name);
            project.SetupGet(p => p.References).Returns(references.Object);
            return project;
        }
    }
}
