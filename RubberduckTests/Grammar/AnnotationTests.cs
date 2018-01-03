using NUnit.Framework;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;

namespace RubberduckTests.Grammar
{
    [TestFixture]
    public class AnnotationTests
    {
        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void TestModuleAnnotation_TypeIsTestModule()
        {
            var annotation = new TestModuleAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.TestModule, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void ModuleInitializeAnnotation_TypeIsModuleInitialize()
        {
            var annotation = new ModuleInitializeAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.ModuleInitialize, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void ModuleCleanupAnnotation_TypeIsModuleCleanup()
        {
            var annotation = new ModuleCleanupAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.ModuleCleanup, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void TestMethodAnnotation_TypeIsTestTest()
        {
            var annotation = new TestMethodAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.TestMethod, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void TestInitializeAnnotation_TypeIsTestInitialize()
        {
            var annotation = new TestInitializeAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.TestInitialize, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void TestCleanupAnnotation_TypeIsTestCleanup()
        {
            var annotation = new TestCleanupAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.TestCleanup, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void IgnoreTestAnnotation_TypeIsIgnoreTest()
        {
            var annotation = new IgnoreTestAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.IgnoreTest, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void IgnoreAnnotation_TypeIsIgnore()
        {
            var annotation = new IgnoreAnnotation(new QualifiedSelection(), new[] { "param" });
            Assert.AreEqual(AnnotationType.Ignore, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void FolderAnnotation_TypeIsFolder()
        {
            var annotation = new FolderAnnotation(new QualifiedSelection(), new[] { "param" });
            Assert.AreEqual(AnnotationType.Folder, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void NoIndentAnnotation_TypeIsNoIndent()
        {
            var annotation = new NoIndentAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.NoIndent, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void InterfaceAnnotation_TypeIsInterface()
        {
            var annotation = new InterfaceAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.Interface, annotation.AnnotationType);
        }

        [Test]
        public void DescriptionAnnotation_TypeIsDescription()
        {
            var annotation = new DescriptionAnnotation(new QualifiedSelection(), null);
            Assert.AreEqual(AnnotationType.Description, annotation.AnnotationType);
        }
    }
}