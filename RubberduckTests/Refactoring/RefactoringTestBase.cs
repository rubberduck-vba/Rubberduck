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
        protected Mock<VBProject> Project;
        protected Mock<VBComponent> Component;
        protected Mock<CodeModule> Module;
        protected Mock<VBE> IDE;

        [TestCleanup]
        public void CleanUp()
        {
            Project = null;
            Component = null;
            Module = null;
        }

        protected QualifiedSelection GetQualifiedSelection(Selection selection)
        {
            return new QualifiedSelection(new QualifiedModuleName(Component.Object), selection);
        }

        protected void SetupProject(string inputCode)
        {
            var window = MockFactory.CreateWindowMock(string.Empty);
            var windows = new MockWindowsCollection(window.Object);

            IDE = MockFactory.CreateVbeMock(windows);

            var codePane = MockFactory.CreateCodePaneMock(IDE, window);

            IDE.SetupGet(vbe => vbe.ActiveCodePane).Returns(codePane.Object);

            Module = MockFactory.CreateCodeModuleMock(inputCode, codePane.Object);

            codePane.SetupGet(p => p.CodeModule).Returns(Module.Object);

            Project = MockFactory.CreateProjectMock("VBAProject", vbext_ProjectProtection.vbext_pp_none);

            Component = MockFactory.CreateComponentMock("Module1", Module.Object, vbext_ComponentType.vbext_ct_StdModule);

            var components = MockFactory.CreateComponentsMock(new List<VBComponent>() { Component.Object });
            components.SetupGet(c => c.Parent).Returns(Project.Object);

            Project.SetupGet(p => p.VBComponents).Returns(components.Object);
            Component.SetupGet(c => c.Collection).Returns(components.Object);
        }
    }
}
