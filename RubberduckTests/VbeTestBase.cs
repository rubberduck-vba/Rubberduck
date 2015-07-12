using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using MockFactory = RubberduckTests.Mocks.MockFactory;

namespace RubberduckTests
{
    public abstract class VbeTestBase
    {
        private Mock<VBE> _ide;
        private ICollection<VBProject> _projects;

        [TestInitialize]
        public void Initialize()
        {
            _ide = MockFactory.CreateVbeMock();

            _projects = new List<VBProject>();
            var projects = MockFactory.CreateProjectsMock(_projects);
            projects.Setup(m => m.Item(It.IsAny<int>())).Returns<int>(i => _projects.ElementAt(i));

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
                _ide.Object.ActiveVBProject = _ide.Object.VBProjects.Item(0);
                _ide.Object.ActiveCodePane = _ide.Object.ActiveVBProject.VBComponents.Item(0).CodeModule.CodePane;
            }
            return new QualifiedSelection(new QualifiedModuleName(_ide.Object.ActiveCodePane.CodeModule.Parent), selection, new RubberduckCodePaneFactory());
        }

        protected QualifiedSelection GetQualifiedSelection(Selection selection, VBComponent component)
        {
            return new QualifiedSelection(new QualifiedModuleName(component), selection, new RubberduckCodePaneFactory());
        }

        protected Mock<VBProject> SetupMockProject(string inputCode, string projectName = null, string moduleName = null, vbext_ComponentType? componentType = null)
        {
            if (componentType == null)
            {
                componentType = vbext_ComponentType.vbext_ct_StdModule;
            }

            if (moduleName == null)
            {
                moduleName = componentType == vbext_ComponentType.vbext_ct_StdModule 
                    ? "Module1" 
                    : componentType == vbext_ComponentType.vbext_ct_ClassModule
                        ? "Class1"
                        : componentType == vbext_ComponentType.vbext_ct_MSForm
                            ? "Form1"
                            : "Document1";
            }

            if (projectName == null)
            {
                projectName = "VBAProject";
            }

            var component = CreateMockComponent(inputCode, moduleName, componentType.Value);
            var components = new List<Mock<VBComponent>> {component};

            var project = CreateMockProject(projectName, vbext_ProjectProtection.vbext_pp_none, components);
            return project;
        }

        protected Mock<VBProject> CreateMockProject(string name, vbext_ProjectProtection protection, ICollection<Mock<VBComponent>> components)
        {
            var project = MockFactory.CreateProjectMock(name, protection);
            var projectComponents = SetupMockComponents(components, project.Object);

            project.SetupGet(m => m.VBE).Returns(_ide.Object);
            project.SetupGet(m => m.VBComponents).Returns(projectComponents.Object);

            _projects.Add(project.Object);
            return project;
        }

        protected Mock<VBComponent> CreateMockComponent(string content, string name, vbext_ComponentType type)
        {
            var module = SetupMockCodeModule(content, name);
            var component = MockFactory.CreateComponentMock(name, module.Object, type, _ide);

            module.SetupGet(m => m.Parent).Returns(component.Object);
            return component;
        }


        private Mock<VBComponents> SetupMockComponents(ICollection<Mock<VBComponent>> items, VBProject project)
        {
            var components = MockFactory.CreateComponentsMock(items, project);
            components.SetupGet(m => m.Parent).Returns(project);
            components.SetupGet(m => m.VBE).Returns(_ide.Object);
            components.Setup(m => m.Item(It.IsAny<int>())).Returns((int index) => items.ElementAt(index).Object);
            components.Setup(m => m.Item(It.IsAny<string>())).Returns((string name) => items.Single(e => e.Object.Name == name).Object);

            return components;
        }

        private Mock<CodeModule> SetupMockCodeModule(string content, string name)
        {
            var codePane = MockFactory.CreateCodePaneMock(_ide, name);
            var module = MockFactory.CreateCodeModuleMock(content, codePane, _ide);

            codePane.SetupGet(m => m.CodeModule).Returns(module.Object);
            return module;
        }
    }
}
