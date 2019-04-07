using Moq;
using NUnit.Framework;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace RubberduckTests.VBEditor
{
    [TestFixture()]
    public class CodePaneSourceHandlerTests
    {
        [Test]
        [Category("COM")]
        public void SourceCodeReturnsContentForModule()
        {
            var codeModuleMock = new Mock<ICodeModule>();
            codeModuleMock.Setup(m => m.Content()).Returns("TestTestTest");

            var codeModule = codeModuleMock.Object;
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var projectsProvider = TestProvider(module, codeModule);
            var codePaneSourceHandler = new CodePaneHandler(projectsProvider);

            var expected = "TestTestTest";
            var actual = codePaneSourceHandler.SourceCode(module);
            
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("COM")]
        public void SourceCodeReturnsEmptyStringForUnrelatedModule()
        {
            var codeModuleMock = new Mock<ICodeModule>();
            codeModuleMock.Setup(m => m.Content()).Returns("TestTestTest");

            var codeModule = codeModuleMock.Object;
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var projectsProvider = TestProvider(module, codeModule);
            var codePaneSourceHandler = new CodePaneHandler(projectsProvider);

            var otherModule = new QualifiedModuleName("TestProject", string.Empty, "OtherTestModule");

            var expected = string.Empty;
            var actual = codePaneSourceHandler.SourceCode(otherModule);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("COM")]
        public void AfterSubstituteCodeTheCorrespondingModuleContainsTheCodeProvided()
        {
            var code = "TestTestTest";
            var codeModuleMock = new Mock<ICodeModule>();
            codeModuleMock.Setup(m => m.Content()).Returns(() => code);
            codeModuleMock.Setup(m => m.CountOfLines).Returns(1);
            codeModuleMock.Setup(m => m.Clear()).Callback(() => code = string.Empty);
            codeModuleMock.Setup(m => m.InsertLines(It.IsAny<int>(), It.IsAny<string>()))
                .Callback((int start, string str) => code = code.Insert(start - 1, str));

            var codeModule = codeModuleMock.Object;
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var projectsProvider = TestProvider(module, codeModule);
            var codePaneSourceHandler = new CodePaneHandler(projectsProvider);

            codePaneSourceHandler.SubstituteCode(module, "NewNewNew");

            var expected = "NewNewNew";
            var actual = codeModule.Content();

            Assert.AreEqual(expected, actual);
        }

        private IProjectsProvider TestProvider(QualifiedModuleName module, ICodeModule testModule)
        {
            var component = new Mock<IVBComponent>();
            component.Setup(c => c.CodeModule).Returns(testModule);
            var provider = new Mock<IProjectsProvider>();
            provider.Setup(p => p.Component(It.IsAny<QualifiedModuleName>()))
                .Returns<QualifiedModuleName>(qmn => qmn.Equals(module) ? component.Object : null);
            return provider.Object;
        }
    }
}
