using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;

namespace RubberduckTests.Grammar
{
    [TestClass]
    public class AnnotationTests
    {
        [TestCategory("Grammar")]
        [TestCategory("Annotations")]
        [TestMethod]
        public void TestModuleAnnotation_TypeIsTestModule()
        {
            var annotation = new TestModuleAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.TestModule, annotation.AnnotationType);
        }

        [TestCategory("Grammar")]
        [TestCategory("Annotations")]
        [TestMethod]
        public void ModuleInitializeAnnotation_TypeIsModuleInitialize()
        {
            var annotation = new ModuleInitializeAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.ModuleInitialize, annotation.AnnotationType);
        }

        [TestCategory("Grammar")]
        [TestCategory("Annotations")]
        [TestMethod]
        public void ModuleCleanupAnnotation_TypeIsModuleCleanup()
        {
            var annotation = new ModuleCleanupAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.ModuleCleanup, annotation.AnnotationType);
        }

        [TestCategory("Grammar")]
        [TestCategory("Annotations")]
        [TestMethod]
        public void TestMethodAnnotation_TypeIsTestTest()
        {
            var annotation = new TestMethodAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.TestMethod, annotation.AnnotationType);
        }

        [TestCategory("Grammar")]
        [TestCategory("Annotations")]
        [TestMethod]
        public void TestInitializeAnnotation_TypeIsTestInitialize()
        {
            var annotation = new TestInitializeAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.TestInitialize, annotation.AnnotationType);
        }

        [TestCategory("Grammar")]
        [TestCategory("Annotations")]
        [TestMethod]
        public void TestCleanupAnnotation_TypeIsTestCleanup()
        {
            var annotation = new TestCleanupAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.TestCleanup, annotation.AnnotationType);
        }

        [TestCategory("Grammar")]
        [TestCategory("Annotations")]
        [TestMethod]
        public void IgnoreTestAnnotation_TypeIsIgnoreTest()
        {
            var annotation = new IgnoreTestAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.IgnoreTest, annotation.AnnotationType);
        }

        [TestCategory("Grammar")]
        [TestCategory("Annotations")]
        [TestMethod]
        public void IgnoreAnnotation_TypeIsIgnore()
        {
            var annotation = new IgnoreAnnotation(new QualifiedSelection(), new[] { "param" });
            Assert.AreEqual(AnnotationType.Ignore, annotation.AnnotationType);
        }

        [TestCategory("Grammar")]
        [TestCategory("Annotations")]
        [TestMethod]
        public void FolderAnnotation_TypeIsFolder()
        {
            var annotation = new FolderAnnotation(new QualifiedSelection(), new[] { "param" });
            Assert.AreEqual(AnnotationType.Folder, annotation.AnnotationType);
        }

        [TestCategory("Grammar")]
        [TestCategory("Annotations")]
        [TestMethod]
        public void NoIndentAnnotation_TypeIsNoIndent()
        {
            var annotation = new NoIndentAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.NoIndent, annotation.AnnotationType);
        }

        [TestCategory("Grammar")]
        [TestCategory("Annotations")]
        [TestMethod]
        public void InterfaceAnnotation_TypeIsInterface()
        {
            var annotation = new InterfaceAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.Interface, annotation.AnnotationType);
        }

        [TestMethod]
        public void DescriptionAnnotation_TypeIsDescription()
        {
            var annotation = new DescriptionAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.Description, annotation.AnnotationType);
        }
    }
}