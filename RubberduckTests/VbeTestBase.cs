using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using MockFactory = RubberduckTests.Mocks.MockFactory;

namespace RubberduckTests
{
    public abstract class VbeTestBase
    {
        private Mock<IVBE> _ide;
        private ICollection<IVBProject> _projects;

        [TestInitialize]
        public void Initialize()
        {
            _ide = MockFactory.CreateVbeMock();

            _projects = new List<IVBProject>();
            var projects = MockFactory.CreateProjectsMock(_projects);
            projects.Setup(m => m[It.IsAny<int>()]).Returns<int>(i => _projects.ElementAt(i));

            _ide.SetupGet(m => m.VBProjects).Returns(() => projects.Object);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _ide = null;
        }

        protected QualifiedSelection GetQualifiedSelection(Selection selection)
        {
            if (_ide.Object.ActiveCodePane == null)
            {
                _ide.Object.ActiveVBProject = _ide.Object.VBProjects[0];
                _ide.Object.ActiveCodePane = _ide.Object.ActiveVBProject.VBComponents[0].CodeModule.CodePane;
            }
            return new QualifiedSelection(new QualifiedModuleName(_ide.Object.ActiveCodePane.CodeModule.Parent), selection);
        }

        protected Mock<IVBProject> SetupMockProject(string inputCode, string projectName = null, string moduleName = null, ComponentType? componentType = null)
        {
            if (componentType == null)
            {
                componentType = ComponentType.StandardModule;
            }

            if (moduleName == null)
            {
                moduleName = componentType == ComponentType.StandardModule 
                    ? "Module1" 
                    : componentType == ComponentType.ClassModule
                        ? "Class1"
                        : componentType == ComponentType.UserForm
                            ? "Form1"
                            : "Document1";
            }

            if (projectName == null)
            {
                projectName = "VBAProject";
            }

            var component = CreateMockComponent(inputCode, moduleName, componentType.Value);
            var components = new List<Mock<IVBComponent>> {component};

            var project = CreateMockProject(projectName, ProjectProtection.Unprotected, components);
            return project;
        }

        protected Mock<IVBProject> CreateMockProject(string name, ProjectProtection protection, ICollection<Mock<IVBComponent>> components)
        {
            var project = MockFactory.CreateProjectMock(name, protection);
            var projectComponents = SetupMockComponents(components, project.Object);

            project.SetupGet(m => m.VBE).Returns(_ide.Object);
            project.SetupGet(m => m.VBComponents).Returns(projectComponents.Object);

            _projects.Add(project.Object);
            return project;
        }

        protected Mock<IVBComponent> CreateMockComponent(string content, string name, ComponentType type)
        {
            var module = SetupMockCodeModule(content, name);
            var component = MockFactory.CreateComponentMock(name, module.Object, type, _ide);

            module.SetupGet(m => m.Parent).Returns(component.Object);
            return component;
        }        private Mock<IVBComponents> SetupMockComponents(ICollection<Mock<IVBComponent>> items, IVBProject project)
        {
            var components = MockFactory.CreateComponentsMock(items, project);
            components.SetupGet(m => m.Parent).Returns(project);
            components.SetupGet(m => m.VBE).Returns(_ide.Object);
            components.Setup(m => m[It.IsAny<int>()]).Returns((int index) => items.ElementAt(index).Object);
            components.Setup(m => m[It.IsAny<string>()]).Returns((string name) => items.Single(e => e.Object.Name == name).Object);

            return components;
        }

        private Mock<ICodeModule> SetupMockCodeModule(string content, string name)
        {
            var codePane = MockFactory.CreateCodePaneMock(_ide, name);
            var module = MockFactory.CreateCodeModuleMock(content, codePane, _ide);
            module.SetupProperty(m => m.Name, name);

            codePane.SetupGet(m => m.CodeModule).Returns(module.Object);
            return module;
        }
    }
}
