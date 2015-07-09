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
        private readonly Mock<Window> _window;

        protected RefactoringTestBase()
        {
            _window = MockFactory.CreateWindowMock(string.Empty);
            var windows = new MockWindowsCollection(_window.Object);

            _ide = MockFactory.CreateVbeMock(windows);
            AttachParentIDE(_window);
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

            var codePane = MockFactory.CreateCodePaneMock(_ide, _window);
            var module = MockFactory.CreateCodeModuleMock(inputCode, codePane.Object);
            var project = MockFactory.CreateProjectMock("VBAProject", vbext_ProjectProtection.vbext_pp_none);
            var component = MockFactory.CreateComponentMock(moduleName, module.Object, componentType.Value);
            var components = MockFactory.CreateComponentsMock(new List<VBComponent>() { component.Object });

            _ide.Setup(m => m.ActiveCodePane).Returns(codePane.Object);
            codePane.Setup(m => m.CodeModule).Returns(module.Object);
            project.Setup(m => m.VBComponents).Returns(components.Object);
            components.Setup(m => m.Parent).Returns(project.Object);
            component.Setup(m => m.Collection).Returns(components.Object);

            AttachParentIDE(module);
            AttachParentIDE(codePane);
            AttachParentIDE(project);
            AttachParentIDE(components);
            AttachParentIDE(component);

            return project;
        }

        private void AttachParentIDE(Mock<Window> mock)
        {
            mock.Setup(m => m.VBE).Returns(_ide.Object);
        }

        private void AttachParentIDE(Mock<CodeModule> mock)
        {
            mock.Setup(m => m.VBE).Returns(_ide.Object);
        }

        private void AttachParentIDE(Mock<CodePane> mock)
        {
            mock.Setup(m => m.VBE).Returns(_ide.Object);
        }

        private void AttachParentIDE(Mock<VBProject> mock)
        {
            mock.Setup(m => m.VBE).Returns(_ide.Object);
        }

        private void AttachParentIDE(Mock<VBComponents> mock)
        {
            mock.Setup(m => m.VBE).Returns(_ide.Object);
        }

        private void AttachParentIDE(Mock<VBComponent> mock)
        {
            mock.Setup(m => m.VBE).Returns(_ide.Object);
        }
    }
}
