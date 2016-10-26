using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Moq;
using Rubberduck.UI;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    static class MockFactory
    {
        /// <summary>
        /// Creates a mock <see cref="IWindow"/> that is particularly useful for passing into <see cref="MockWindowsCollection"/>'s ctor.
        /// </summary>
        /// <returns>
        /// A <see cref="Mock{IWindow}"/>that has all the properties needed for <see cref="DockableToolwindowPresenter"/> pre-setup.
        /// </returns>
        internal static Mock<IWindow> CreateWindowMock()
        {
            var window = new Mock<IWindow>();
            window.SetupProperty(w => w.IsVisible, false);
            window.SetupGet(w => w.LinkedWindows).Returns((ILinkedWindows) null);
            window.SetupProperty(w => w.Height);
            window.SetupProperty(w => w.Width);

            return window;
        }

        /// <summary>
        /// Creates a mock <see cref="IWindow"/> with it's <see cref="IWindow.Caption"/> propery set up.
        /// </summary>
        /// <param name="caption">The value to return from <see cref="IWindow.Caption"/>.</param>
        /// <returns>
        /// A <see cref="Mock{Window}"/>that has all the properties needed for <see cref="DockableToolwindowPresenter"/> pre-setup.
        /// </returns>
        internal static Mock<IWindow> CreateWindowMock(string caption)
        {
            var window = CreateWindowMock();
            window.SetupGet(w => w.Caption).Returns(caption);

            return window;
        }

        internal static Mock<IVBE> CreateVbeMock()
        {
            return CreateVbeMock(new MockWindowsCollection());
        }

        /// <summary>
        /// Creates a new <see cref="Mock{VBE}"/> that returns the <see cref="IWindows"/> collection argument out of the Windows property.
        /// </summary>
        /// <param name="windows">
        /// A <see cref="MockWindowsCollection"/> is expected. 
        /// Other objects implementing the<see cref="IWindows"/> interface could cause issues.
        /// </param>
        /// <returns></returns>
        internal static Mock<IVBE> CreateVbeMock(MockWindowsCollection windows)
        {
            var vbe = new Mock<IVBE>();
            windows.VBE = vbe.Object;
            vbe.Setup(m => m.Windows).Returns(windows);
            vbe.SetupProperty(m => m.ActiveCodePane);
            vbe.SetupProperty(m => m.ActiveVBProject);
            vbe.SetupGet(m => m.SelectedVBComponent).Returns(() => vbe.Object.ActiveCodePane.CodeModule.Parent);
            vbe.SetupGet(m => m.ActiveWindow).Returns(() => vbe.Object.ActiveCodePane.Window);

            //setting up a main window lets the native window functions fun
            var mainWindow = new Mock<IWindow>();
            mainWindow.Setup(m => m.HWnd).Returns(0);

            vbe.SetupGet(m => m.MainWindow).Returns(mainWindow.Object);

            return vbe;
        }

        /// <summary>
        /// Creates a "selectable" <see cref="Mock{ICodePane}"/>.
        /// </summary>
        /// <param name="vbe">Returned back from the <see cref="ICodePane.VBE"/> property.</param>
        /// <param name="name">The caption of the window object that will be created for this code pane.</param>
        /// <returns></returns>
        internal static Mock<ICodePane> CreateCodePaneMock(Mock<IVBE> vbe, string name)
        {
            var windows = vbe.Object.Windows as MockWindowsCollection;
            if (windows == null)
            {
                return null;
            }

            var codePane = new Mock<ICodePane>();
            var window = windows.CreateWindow(name);
            windows.Add(window);

            codePane.Setup(p => p.SetSelection(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<int>(), It.IsAny<int>()));
            codePane.Setup(p => p.Show());
            codePane.SetupGet(p => p.VBE).Returns(vbe.Object);
            codePane.SetupGet(p => p.Window).Returns(window);
            return codePane;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{ICodeModule}"/> setup to appropriately mimic getting and modifying code contained in the <see cref="ICodeModule"/>.
        /// </summary>
        /// <param name="code">A block of VBA code.</param>
        /// <returns></returns>
        private static Mock<ICodeModule> CreateCodeModuleMock(string code)
        {
            var lines = code.Split(new[] {Environment.NewLine}, StringSplitOptions.None).ToList();

            var codeModule = new Mock<ICodeModule>();
            codeModule.SetupGet(c => c.CountOfLines).Returns(lines.Count);

            // ReSharper disable once UseIndexedProperty
            // No R#, the indexed property breaks the expression. I tried that first.
            codeModule.Setup(m => m.GetLines(It.IsAny<int>(), It.IsAny<int>()))
                .Returns<int, int>((start, count) => string.Join(Environment.NewLine, lines.Skip(start - 1).Take(count)));

            codeModule.Setup(m => m.ReplaceLine(It.IsAny<int>(), It.IsAny<string>()))
                .Callback<int, string>((index, str) => lines[index - 1] = str);

            codeModule.Setup(m => m.DeleteLines(It.IsAny<int>(), It.IsAny<int>()))
                .Callback<int, int>((index, count) => lines.RemoveRange(index - 1, count));

            codeModule.Setup(m => m.InsertLines(It.IsAny<int>(), It.IsAny<string>()))
                .Callback<int, string>((index, newLine) => lines.Insert(index - 1, newLine));
                
            return codeModule;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{ICodeModule}"/> setup to appropriately mimic getting and modifying code contained in the <see cref="ICodeModule"/>.
        /// </summary>
        /// <param name="code">A block of VBA code.</param>
        /// <param name="codePane">Returned back from the <see cref="ICodeModule.CodePane"/> property.</param>
        /// <returns></returns>
        internal static Mock<ICodeModule> CreateCodeModuleMock(string code, Mock<ICodePane> codePane, Mock<IVBE> vbe)
        {
            var codeModule = CreateCodeModuleMock(code);
            codeModule.SetupGet(m => m.CodePane).Returns(codePane.Object);
            codeModule.SetupGet(m => m.VBE).Returns(vbe.Object);

            return codeModule;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{VBComponent}"/>.
        /// </summary>
        /// <param name="name">The name to return from the <see cref="IVBComponent.Name"/> property.</param>
        /// <param name="codeModule">The <see cref="ICodeModule"/> to return from the CodeModule property.</param>
        /// <param name="componentType">
        /// The type of component to be simulated.
        /// Use vbext_ct_StdModule for standard modules.
        /// Use vbext_ct_ClassModule for classes.
        /// vbext_ct_ActiveXDesigner is invalid for the VBE.
        /// </param>
        /// <returns></returns>
        internal static Mock<IVBComponent> CreateComponentMock(string name, ICodeModule codeModule, ComponentType componentType, Mock<IVBE> vbe)
        {
            var component = new Mock<IVBComponent>();
            component.SetupProperty(m => m.Name, name);
            component.SetupGet(m => m.CodeModule).Returns(codeModule);
            component.SetupGet(m => m.Type).Returns(componentType);
            component.SetupGet(m => m.VBE).Returns(vbe.Object);
            return component;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{IVBComponents}"/> that can be iterated over as an <see cref="IEnumerable"/>.
        /// </summary>
        /// <param name="componentList">The collection to be iterated over.</param>
        /// <returns></returns>
        internal static Mock<IVBComponents> CreateComponentsMock(IEnumerable<IVBComponent> componentList)
        {
            var components = new Mock<IVBComponents>();
            components.Setup(c => c.GetEnumerator()).Returns(componentList.GetEnumerator());
            components.As<IEnumerable>().Setup(c => c.GetEnumerator()).Returns(componentList.GetEnumerator());

            return components;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{IVBComponents}"/> that can be iterated over as an <see cref="IEnumerable"/>.
        /// </summary>
        /// <param name="componentList">The collection to be iterated over.</param>
        /// <param name="project">The <see cref="IVBComponents.Parent"/> property.</param>
        /// <returns></returns>
        internal static Mock<IVBComponents> CreateComponentsMock(ICollection<Mock<IVBComponent>> componentList, IVBProject project)
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
        /// Creates a new <see cref="Mock{IVBProject}"/>.
        /// </summary>
        /// <param name="name">The <see cref="IVBProject.Name"/> property.</param>
        /// <param name="protectionLevel">
        /// The <see cref="IVBProject.Protection"/> property.
        /// Use vbext_pp_none to simulate a normal project.
        /// Use vbext_pp_locked to simulate a password protected, or otherwise unavailable, project.
        /// </param>
        /// <returns></returns>
        internal static Mock<IVBProject> CreateProjectMock(string name, ProjectProtection protectionLevel)
        {
            var project = new Mock<IVBProject>();
            project.SetupProperty(p => p.Name, name);
            project.SetupGet(p => p.Protection).Returns(protectionLevel);
            return project;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{IVBProjects}"/> that can be iterated over as an <see cref="IEnumerable"/>.
        /// </summary>
        /// <param name="projectList">The collection to be iterated over.</param>
        /// <returns></returns>
        internal static Mock<IVBProjects> CreateProjectsMock(ICollection<IVBProject> projectList)
        {
            var projects = new Mock<IVBProjects>();
            projects.Setup(p => p.GetEnumerator()).Returns(projectList.GetEnumerator());
            projects.As<IEnumerable>().Setup(p => p.GetEnumerator()).Returns(projectList.GetEnumerator());

            return projects;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{IReference}"/>.
        /// </summary>
        /// <param name="name">The see<see cref="IReference.Name"/>.</param>
        /// <param name="filePath">The <see cref="IReference.FullPath"/> filepath.</param>
        /// <returns></returns>
        internal static Mock<IReference> CreateMockReference(string name, string filePath)
        {
            var reference = new Mock<IReference>();
            reference.SetupGet(r => r.Name).Returns(name);
            reference.SetupGet(r => r.FullPath).Returns(filePath);

            return reference;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{References}"/> collection that can be iterated over as an <see cref="IEnumerable"/>.
        /// </summary>
        /// <param name="referenceList">The collection to be iterated over.</param>
        /// <returns></returns>
        internal static Mock<IReferences> CreateReferencesMock(List<IReference> referenceList)
        {
            var references = new Mock<IReferences>();
            references.Setup(r => r.GetEnumerator()).Returns(referenceList.GetEnumerator());
            references.As<IEnumerable>().Setup(r => r.GetEnumerator()).Returns(referenceList.GetEnumerator());
            return references;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{IProject}"/> that is set up with a <see cref="IReferences"/> collection.
        /// </summary>
        /// <param name="name">The <see cref="IVBProject"/> <see cref="IVBProject.Name"/>.</param>
        /// <param name="references">The <see cref="IReferences"/> collection.</param>
        /// <returns></returns>
        internal static Mock<IVBProject> CreateProjectMock(string name, Mock<IReferences> references)
        {
            var project = new Mock<IVBProject>();
            project.SetupProperty(p => p.Name, name);
            project.SetupGet(p => p.References).Returns(references.Object);
            return project;
        }
    }
}
