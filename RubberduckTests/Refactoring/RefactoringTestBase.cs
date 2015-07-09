using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;
using MockFactory = RubberduckTests.Mocks.MockFactory;

namespace RubberduckTests.Refactoring
{
    public abstract class RefactoringTestBase
    {
        private readonly Mock<VBE> _ide;

        protected RefactoringTestBase()
        {
            _ide = MockFactory.CreateVbeMock();
        }

        protected QualifiedSelection GetQualifiedSelection(Selection selection)
        {
            return GetQualifiedSelection(selection, _ide.Object.ActiveVBProject.VBComponents.Item(0));
        }

        protected QualifiedSelection GetQualifiedSelection(Selection selection, VBComponent component)
        {
            return new QualifiedSelection(new QualifiedModuleName(component), selection);
        }

        protected Mock<VBProject> SetupMockProject(string inputCode, string moduleName = null, vbext_ComponentType? componentType = null)
        {
            if (componentType == null)
            {
                componentType = vbext_ComponentType.vbext_ct_StdModule;
            }

            if (moduleName == null)
            {
                moduleName = "Module1";
            }

            var project = MockFactory.CreateProjectMock("VBAProject", vbext_ProjectProtection.vbext_pp_none);
            var component = CreateMockComponent(inputCode, moduleName, componentType.Value);
            var components = SetupMockComponents(new List<Mock<VBComponent>> {component}, project.Object);

            _ide.SetupGet(m => m.ActiveCodePane).Returns(component.Object.CodeModule.CodePane);
            _ide.SetupSet(vbe => vbe.ActiveCodePane = It.IsAny<CodePane>()); // todo: verify that this works as expected
            
            _ide.SetupGet(m => m.ActiveVBProject).Returns(project.Object);
            _ide.SetupSet(vbe => vbe.ActiveVBProject = It.IsAny<VBProject>()); // todo: verify that this works as expected

            _ide.SetupGet(m => m.ActiveWindow).Returns(_ide.Object.ActiveCodePane.Window);

            project.SetupGet(m => m.VBComponents).Returns(components.Object);
            components.Setup(m => m.Item(0)).Returns(component.Object);
            components.SetupGet(m => m.Parent).Returns(project.Object);
            component.SetupGet(m => m.Collection).Returns(components.Object);

            project.SetupGet(m => m.VBE).Returns(_ide.Object);

            return project;
        }

        public Mock<VBComponent> CreateMockComponent(string content, string name, vbext_ComponentType type)
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
