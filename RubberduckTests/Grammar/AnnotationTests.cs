using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;

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
        public void InterfaceAnnotation_TypeIsInterface()
        {
            var annotation = new InterfaceAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.Interface, annotation.AnnotationType);
        }
    }
}