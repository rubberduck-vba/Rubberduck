using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Annotations
{
    [TestFixture]
    public class AttributeAnnotationTests
    {
        [TestCase("VB_Description", "\"SomeDescription\"", "ModuleDescription", "\"SomeDescription\"")]
        [TestCase("VB_Exposed", "True", "Exposed")]
        [TestCase("VB_PredeclaredId", "True", "PredeclaredId")]
        public void ModuleAttributeAnnotationReturnsReturnsCorrectAttribute(string expectedAttribute, string expectedAttributeValues, string annotationName, string annotationValue = null)
        {
            var code = $@"
'@{annotationName} {annotationValue}";

            var vbe = MockVbeBuilder.BuildFromSingleModule(code, "Class1", ComponentType.ClassModule, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var moduleDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.ClassModule).Single();
                var moduleAnnotations = moduleDeclaration.Annotations
                    .Select(pta => (pta.Annotation, pta.AnnotationArguments));
                var (annotation, annotationArguments) = moduleAnnotations.Single(tpl => tpl.Annotation.Name.Equals(annotationName));

                var actualAttribute = ((IAttributeAnnotation) annotation).Attribute(annotationArguments);
                var actualAttributeValues = ((IAttributeAnnotation)annotation).AnnotationToAttributeValues(annotationArguments);
                var actualAttributesValuesText = string.Join(", ", actualAttributeValues);

                Assert.AreEqual(expectedAttribute, actualAttribute);
                Assert.AreEqual(actualAttributesValuesText, expectedAttributeValues);
            }
        }

        [TestCase("VB_ProcData.VB_Invoke_Func", @"""A\n14""", "ExcelHotkey", "A")] //See issue #5268 at https://github.com/rubberduck-vba/Rubberduck/issues/5268
        [TestCase("VB_Description", "\"SomeDescription\"", "Description", "\"SomeDescription\"")]
        [TestCase("VB_UserMemId", "0", "DefaultMember")]
        [TestCase("VB_UserMemId", "-4", "Enumerator")]
        public void MemberAttributeAnnotationReturnsReturnsCorrectAttribute(string expectedAttribute, string expectedAttributeValues, string annotationName, string annotationValue = null)
        {
            var code = $@"
'@{annotationName} {annotationValue}
Public Function Foo()
End Function";

            var vbe = MockVbeBuilder.BuildFromSingleModule(code, "Class1", ComponentType.ClassModule, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var memberDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function).Single();
                var memberAnnotations = memberDeclaration.Annotations
                    .Select(pta => (pta.Annotation, pta.AnnotationArguments));
                var (annotation, annotationArguments) = memberAnnotations.Single(tpl => tpl.Annotation.Name.Equals(annotationName));

                var actualAttribute = ((IAttributeAnnotation)annotation).Attribute(annotationArguments);
                var actualAttributeValues = ((IAttributeAnnotation) annotation).AnnotationToAttributeValues(annotationArguments);
                var actualAttributesValuesText = string.Join(", ", actualAttributeValues);

                Assert.AreEqual(expectedAttribute, actualAttribute);
                Assert.AreEqual(expectedAttributeValues, actualAttributesValuesText);
            }
        }

        [TestCase("VB_VarDescription", "\"SomeDescription\"", "VariableDescription", "\"SomeDescription\"")]
        public void VariableAttributeAnnotationReturnsReturnsCorrectAttribute(string expectedAttribute, string expectedAttributeValues, string annotationName, string annotationValue = null)
        {
            var code = $@"
'@{annotationName} {annotationValue}
Public MyVariable";

            var vbe = MockVbeBuilder.BuildFromSingleModule(code, "Class1", ComponentType.ClassModule, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var memberDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();
                var memberAnnotations = memberDeclaration.Annotations
                    .Select(pta => (pta.Annotation, pta.AnnotationArguments));
                var (annotation, annotationArguments) = memberAnnotations.Single(tpl => tpl.Annotation.Name.Equals(annotationName));

                var actualAttribute = ((IAttributeAnnotation)annotation).Attribute(annotationArguments);
                var actualAttributeValues = ((IAttributeAnnotation)annotation).AnnotationToAttributeValues(annotationArguments);
                var actualAttributesValuesText = string.Join(", ", actualAttributeValues);

                Assert.AreEqual(expectedAttribute, actualAttribute);
                Assert.AreEqual(actualAttributesValuesText, expectedAttributeValues);
            }
        }
    }
}