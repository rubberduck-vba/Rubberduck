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
        public void NotRecognizedAnnotation_TypeIsNotRecognized()
        {
            var annotation = new NotRecognizedAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.NotRecognized, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void TestModuleAnnotation_TypeIsTestModule()
        {
            var annotation = new TestModuleAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.TestModule, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void ModuleInitializeAnnotation_TypeIsModuleInitialize()
        {
            var annotation = new ModuleInitializeAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.ModuleInitialize, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void ModuleCleanupAnnotation_TypeIsModuleCleanup()
        {
            var annotation = new ModuleCleanupAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.ModuleCleanup, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void TestMethodAnnotation_TypeIsTestTest()
        {
            var annotation = new TestMethodAnnotation(new QualifiedSelection(), null, new[] { "param" });
            Assert.AreEqual(AnnotationType.TestMethod, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void TestInitializeAnnotation_TypeIsTestInitialize()
        {
            var annotation = new TestInitializeAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.TestInitialize, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void TestCleanupAnnotation_TypeIsTestCleanup()
        {
            var annotation = new TestCleanupAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.TestCleanup, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void IgnoreTestAnnotation_TypeIsIgnoreTest()
        {
            var annotation = new IgnoreTestAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.IgnoreTest, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void IgnoreAnnotation_TypeIsIgnore()
        {
            var annotation = new IgnoreAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.Ignore, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void FolderAnnotation_TypeIsFolder()
        {
            var annotation = new FolderAnnotation(new QualifiedSelection(), null, new[] { "param" });
            Assert.AreEqual(AnnotationType.Folder, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void NoIndentAnnotation_TypeIsNoIndent()
        {
            var annotation = new NoIndentAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.NoIndent, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void InterfaceAnnotation_TypeIsInterface()
        {
            var annotation = new InterfaceAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.Interface, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void ModuleAttributeAnnotation_TypeIsModuleAttribute()
        {
            var annotation = new ModuleAttributeAnnotation(new QualifiedSelection(), null, new[] { "Attribute", "Value" });
            Assert.AreEqual(AnnotationType.ModuleAttribute, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void MemberAttributeAnnotation_TypeIsMemberAttribute()
        {
            var annotation = new MemberAttributeAnnotation(new QualifiedSelection(), null, new[] { "Attribute", "Value" });
            Assert.AreEqual(AnnotationType.MemberAttribute, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void DescriptionAnnotation_TypeIsDescription()
        {
            var annotation = new DescriptionAnnotation(new QualifiedSelection(), null, new[] { "Desc"});
            Assert.AreEqual(AnnotationType.Description, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void ModuleDescriptionAnnotation_TypeIsModuleDescription()
        {
            var annotation = new ModuleDescriptionAnnotation(new QualifiedSelection(), null, new[] { "Desc" });
            Assert.AreEqual(AnnotationType.ModuleDescription, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void VariableDescriptionAnnotation_TypeIsModuleDescription()
        {
            var annotation = new VariableDescriptionAnnotation(new QualifiedSelection(), null, new[] { "Desc" });
            Assert.AreEqual(AnnotationType.VariableDescription, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void DefaultMemberAnnotation_TypeIsDefaultMember()
        {
            var annotation = new DefaultMemberAnnotation(new QualifiedSelection(), null, new[] { "param" });
            Assert.AreEqual(AnnotationType.DefaultMember, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void EnumerationMemberAnnotation_TypeIsEnumerator()
        {
            var annotation = new EnumeratorMemberAnnotation(new QualifiedSelection(), null, new[] { "param" });
            Assert.AreEqual(AnnotationType.Enumerator, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void ExposedModuleAnnotation_TypeIsExposed()
        {
            var annotation = new ExposedModuleAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.Exposed, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void IgnoreModuleAnnotation_TypeIsIgnoreModule()
        {
            var annotation = new IgnoreModuleAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.IgnoreModule, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void PredeclaredIdAnnotation_TypeIsPredeclaredId()
        {
            var annotation = new PredeclaredIdAnnotation(new QualifiedSelection(), null, null);
            Assert.AreEqual(AnnotationType.PredeclaredId, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void ObsoleteAnnotation_TypeIsObsolete()
        {
            var annotation = new ObsoleteAnnotation(new QualifiedSelection(), null, new[] { "param" });
            Assert.AreEqual(AnnotationType.Obsolete, annotation.AnnotationType);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void IgnoreAnnotation_CanBeAppliedMultipleTimes()
        {
            var annotation = new IgnoreAnnotation(new QualifiedSelection(), null, null);
            Assert.True(annotation.AllowMultiple);
        }

        [Category("Grammar")]
        [Category("Annotations")]
        [Test]
        public void IgnoreModuleAnnotation_CanBeAppliedMultipleTimes()
        {
            var annotation = new IgnoreModuleAnnotation(new QualifiedSelection(), null, null);
            Assert.True(annotation.AllowMultiple);
        }
    }
}