using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.Grammar
{
    [TestClass]
    public class AnnotationTests
    {
        [TestMethod]
        public void TestModuleAnnotation_TypeIsTestModule()
        {
            var annotation = new TestModuleAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.TestModule, annotation.AnnotationType);
        }

        [TestMethod]
        public void ModuleInitializeAnnotation_TypeIsModuleInitialize()
        {
            var annotation = new ModuleInitializeAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.ModuleInitialize, annotation.AnnotationType);
        }

        [TestMethod]
        public void ModuleCleanupAnnotation_TypeIsModuleCleanup()
        {
            var annotation = new ModuleCleanupAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.ModuleCleanup, annotation.AnnotationType);
        }

        [TestMethod]
        public void TestMethodAnnotation_TypeIsTestTest()
        {
            var annotation = new TestMethodAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.TestMethod, annotation.AnnotationType);
        }

        [TestMethod]
        public void TestInitializeAnnotation_TypeIsTestInitialize()
        {
            var annotation = new TestInitializeAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.TestInitialize, annotation.AnnotationType);
        }

        [TestMethod]
        public void TestCleanupAnnotation_TypeIsTestCleanup()
        {
            var annotation = new TestCleanupAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.TestCleanup, annotation.AnnotationType);
        }

        [TestMethod]
        public void IgnoreTestAnnotation_TypeIsIgnoreTest()
        {
            var annotation = new IgnoreTestAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.IgnoreTest, annotation.AnnotationType);
        }

        [TestMethod, ExpectedException(typeof(InvalidAnnotationArgumentException))]
        public void IgnoreAnnotation_TypeIsIgnore_NoParam()
        {
            var annotation = new IgnoreAnnotation(new QualifiedSelection(), new List<string>());
            Assert.AreEqual(AnnotationType.Ignore, annotation.AnnotationType);
        }

        [TestMethod, ExpectedException(typeof(InvalidAnnotationArgumentException))]
        public void FolderAnnotation_TypeIsFolder_NoParam()
        {
            var annotation = new FolderAnnotation(new QualifiedSelection(), new List<string>());
            Assert.AreEqual(AnnotationType.Folder, annotation.AnnotationType);
        }

        [TestMethod]
        public void IgnoreAnnotation_TypeIsIgnore()
        {
            var annotation = new IgnoreAnnotation(new QualifiedSelection(), new[] { "param" });
            Assert.AreEqual(AnnotationType.Ignore, annotation.AnnotationType);
        }

        [TestMethod]
        public void FolderAnnotation_TypeIsFolder()
        {
            var annotation = new FolderAnnotation(new QualifiedSelection(), new[] { "param" });
            Assert.AreEqual(AnnotationType.Folder, annotation.AnnotationType);
        }

        [TestMethod]
        public void NoIndentAnnotation_TypeIsNoIndent()
        {
            var annotation = new NoIndentAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.NoIndent, annotation.AnnotationType);
        }

        [TestMethod]
        public void DeclarationHasMultipleAnnotations()
        {
            var input =
@"'@TestMethod
'@IgnoreTest
Public Sub Foo()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, input);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var declaration = parser.State.AllUserDeclarations.First(f => f.DeclarationType == DeclarationType.Procedure);

            Assert.IsTrue(declaration.Annotations.Count() == 2);
            Assert.IsTrue(declaration.Annotations.Any(a => a.AnnotationType == AnnotationType.TestMethod));
            Assert.IsTrue(declaration.Annotations.Any(a => a.AnnotationType == AnnotationType.IgnoreTest));
        }
    }
}