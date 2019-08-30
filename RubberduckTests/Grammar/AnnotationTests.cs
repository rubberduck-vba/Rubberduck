using NUnit.Framework;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;

namespace RubberduckTests.Grammar
{
    [TestFixture]
    [Category("Grammar")]
    [Category("Annotations")]
    public class AnnotationTests
    {
        [TestCase(typeof(DefaultMemberAnnotation), "DefaultMember")]
        [TestCase(typeof(DescriptionAnnotation), "Description")]
        [TestCase(typeof(EnumeratorMemberAnnotation), "Enumerator")]
        [TestCase(typeof(ExcelHotKeyAnnotation), "ExcelHotkey")]
        [TestCase(typeof(ExposedModuleAnnotation), "Exposed")]
        [TestCase(typeof(FolderAnnotation), "Folder")]
        [TestCase(typeof(IgnoreAnnotation), "Ignore")]
        [TestCase(typeof(IgnoreModuleAnnotation), "IgnoreModule")]
        [TestCase(typeof(IgnoreTestAnnotation), "IgnoreTest")]
        [TestCase(typeof(InterfaceAnnotation), "Interface")]
        [TestCase(typeof(MemberAttributeAnnotation), "MemberAttribute")]
        [TestCase(typeof(ModuleAttributeAnnotation), "ModuleAttribute")]
        [TestCase(typeof(ModuleCleanupAnnotation), "ModuleCleanup")]
        [TestCase(typeof(ModuleDescriptionAnnotation), "ModuleDescription")]
        [TestCase(typeof(ModuleInitializeAnnotation), "ModuleInitialize")]
        [TestCase(typeof(NoIndentAnnotation), "NoIndent")]
        [TestCase(typeof(NotRecognizedAnnotation), "NotRecognized")]
        [TestCase(typeof(ObsoleteAnnotation), "Obsolete")]
        [TestCase(typeof(PredeclaredIdAnnotation), "PredeclaredId")]
        [TestCase(typeof(TestCleanupAnnotation), "TestCleanup")]
        [TestCase(typeof(TestInitializeAnnotation), "TestInitialize")]
        [TestCase(typeof(TestMethodAnnotation), "TestMethod")]
        [TestCase(typeof(TestModuleAnnotation), "TestModule")]
        [TestCase(typeof(VariableDescriptionAnnotation), "VariableDescription")]
        public void AnnotationTypes_MatchExpectedAnnotationNames(Type annotationType, string expectedName)
        {
            IAnnotation annotation = (IAnnotation)Activator.CreateInstance(annotationType);
            Assert.AreEqual(expectedName, annotation.Name);
        }

        [TestCase(typeof(IgnoreAnnotation))]
        [TestCase(typeof(IgnoreModuleAnnotation))]
        public void AnnotationTypes_MultipleApplicationsAllowed(Type annotationType)
        {
            IAnnotation annotation = (IAnnotation)Activator.CreateInstance(annotationType);
            Assert.IsTrue(annotation.AllowMultiple);
        }
    }
}