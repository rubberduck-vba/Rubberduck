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
        [TestCase(typeof(DefaultMemberAnnotation), "DefaultMember", new[] { "param" })]
        [TestCase(typeof(DescriptionAnnotation), "Description", new[] { "desc" })]
        [TestCase(typeof(EnumeratorMemberAnnotation), "Enumerator", new[] { "param" })]
        [TestCase(typeof(ExcelHotKeyAnnotation), "ExcelHotkey", new [] { "A" })]
        [TestCase(typeof(ExposedModuleAnnotation), "Exposed")]
        [TestCase(typeof(FolderAnnotation), "Folder", new[] { "param" })]
        [TestCase(typeof(IgnoreAnnotation), "Ignore")]
        [TestCase(typeof(IgnoreModuleAnnotation), "IgnoreModule")]
        [TestCase(typeof(IgnoreTestAnnotation), "IgnoreTest")]
        [TestCase(typeof(InterfaceAnnotation), "Interface")]
        [TestCase(typeof(MemberAttributeAnnotation), "MemberAttribute", new[] { "Attribute", "Value" })]
        [TestCase(typeof(ModuleAttributeAnnotation), "ModuleAttribute", new[] { "Attribute", "Value" })]
        [TestCase(typeof(ModuleCleanupAnnotation), "ModuleCleanup")]
        [TestCase(typeof(ModuleDescriptionAnnotation), "ModuleDescription", new[] { "desc" })]
        [TestCase(typeof(ModuleInitializeAnnotation), "ModuleInitialize")]
        [TestCase(typeof(NoIndentAnnotation), "NoIndent")]
        [TestCase(typeof(NotRecognizedAnnotation), "NotRecognized")]
        [TestCase(typeof(ObsoleteAnnotation), "Obsolete", new [] { "justification" })]
        [TestCase(typeof(PredeclaredIdAnnotation), "PredeclaredId")]
        [TestCase(typeof(TestCleanupAnnotation), "TestCleanup")]
        [TestCase(typeof(TestInitializeAnnotation), "TestInitialize")]
        [TestCase(typeof(TestMethodAnnotation), "TestMethod")]
        [TestCase(typeof(TestModuleAnnotation), "TestModule")]
        [TestCase(typeof(VariableDescriptionAnnotation), "VariableDescription", new[] { "desc" })]
        public void AnnotationTypes_MatchExpectedAnnotationNames(Type annotationType, string name, IEnumerable<string> args = null)
        {
            var annotation = (IAnnotation) Activator.CreateInstance(annotationType, new QualifiedSelection(), null, args ?? new List<string>());
            Assert.AreEqual(name, annotation.AnnotationType);
        }

        [TestCase(typeof(IgnoreAnnotation))]
        [TestCase(typeof(IgnoreModuleAnnotation))]
        public void AnnotationTypes_MultipleApplicationsAllowed(Type annotationType)
        {
            var annotation = (IAnnotation)Activator.CreateInstance(annotationType, new QualifiedSelection(), null, null);
            Assert.IsTrue(annotation.MetaInformation.AllowMultiple);
        }
    }
}