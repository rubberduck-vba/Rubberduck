using System;
using System.Collections;
using System.Collections.Generic;
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
        /// Creates a new <see cref="Mock{VBE}"/> that returns the <see cref="Windows"/> collection argument out of the Windows property.
        /// </summary>
        /// <param name="windows">
        /// A <see cref="MockWindowsCollection"/> is expected. 
        /// Other objects implementing the<see cref="Windows"/> interface could cause issues.
        /// </param>
        /// <returns></returns>
        internal static Mock<VBE> CreateVbeMock(Windows windows)
        {
            var vbe = new Mock<VBE>();
            vbe.Setup(v => v.Windows).Returns(windows);

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
        internal static Mock<VBE> CreateVbeMock(Windows windows, VBProjects projects)
        {
            var vbe = CreateVbeMock(windows);
            vbe.SetupGet(v => v.VBProjects).Returns(projects);

            return vbe;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{CodeModule}"/> with <see cref="CodeModule.get_Lines"/> and <see cref="CodeModule.CountOfLines"/> 
        /// setup to appropriately mimic getting code out of the <see cref="CodeModule"/>.
        /// </summary>
        /// <param name="code">A block of VBA code.</param>
        /// <returns></returns>
        internal static Mock<CodeModule> CreateCodeModuleMock(string code)
        {
            var lineCount = code.Split(new [] { Environment.NewLine }, StringSplitOptions.None).Length;

            var codeModule = new Mock<CodeModule>();
            codeModule.SetupGet(c => c.CountOfLines).Returns(lineCount);

            // ReSharper disable once UseIndexedProperty
            // No R#, the indexed property breaks the expression. I tried that first.
            codeModule.SetupGet(c => c.get_Lines(1, lineCount)).Returns(code);
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
        internal static Mock<VBComponent> CreateComponentMock(string name, CodeModule codeModule, vbext_ComponentType componentType)
        {
            var component = new Mock<VBComponent>();
            component.SetupProperty(c => c.Name, name);
            component.SetupGet(c => c.CodeModule).Returns(codeModule);
            component.SetupGet(c => c.Type).Returns(componentType);
            return component;
        }

        /// <summary>
        /// Creates a new <see cref="Mock{VBComponents}"/> that can be iterated over as an <see cref="IEnumerable"/>.
        /// </summary>
        /// <param name="componentList">The collection to be iterated over.</param>
        /// <returns></returns>
        internal static Mock<VBComponents> CreateComponentsMock(List<VBComponent> componentList)
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
        internal static Mock<VBComponents> CreateComponentsMock(List<VBComponent> componentList, VBProject project)
        {
            var components = CreateComponentsMock(componentList);
            components.SetupGet(c => c.Parent).Returns(project);

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
        internal static Mock<VBProjects> CreateProjectsMock(List<VBProject> projectList)
        {
            var projects = new Mock<VBProjects>();
            projects.Setup(p => p.GetEnumerator()).Returns(projectList.GetEnumerator());
            projects.As<IEnumerable>().Setup(p => p.GetEnumerator()).Returns(projectList.GetEnumerator());

            return projects;
        }

        //internal static Mock<VBProjects> CreateProjectsMock(List<VBProject> projectList, VBProject project, VBComponents components)
        //{
        //    CreateProjectsMock(projectList, project);
        //    project.SetupGet(p => p.VBComponents).Returns(components.Object);
        //    return projects;
        //}

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
