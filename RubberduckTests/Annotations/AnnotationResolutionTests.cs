using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using RubberduckTests.Mocks;

namespace RubberduckTests.Annotations
{
    [TestFixture]
    public class AnnotationResolutionTests
    {
        [Test]
        public void MemberAnnotationsAboveMemberGetScopedToMember()
        {
            const string inputCode =
                @"
'@TestMethod
'@Enumerator 17, 12 @DefaultMember
Public Sub Foo()
End Sub

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Procedure).Single();

                var expectedAnnotationCount = 3;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void NonMemberAnnotationsAboveMemberDoNotGetScopedToMember()
        {
            const string inputCode =
                @"
'@TestModule
Public Sub Foo()
End Sub

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Procedure).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void LineContinuedMemberAnnotationsAboveMemberGetScopedToMember()
        {
            const string inputCode =
                @"
Public Sub Foo()
End Sub
'@TestMethod _

'@Enumerator _
17 _
, _
12 _
 _
@DefaultMember _

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function).Single();

                var expectedAnnotationCount = 3;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void MemberAnnotationsAboveFirstNonAnnotationLineAboveMemberStillGetScopedToMember()
        {
            const string inputCode =
                @"
Public Sub Foo()
End Sub
'@TestMethod
'SomeComment
'@Enumerator _
17 _
, _
12 _
 _
@DefaultMember _

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function).Single();

                var expectedAnnotationCount = 3;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void MemberAnnotationsOnInOrBelowMemberDoNotGetScopedToMember()
        {
            const string inputCode =
                @"
Public Sub Foo()
End Sub
'SomeComment
'@Enumerator _
17 _
, _
12 _
 _
@DefaultMember _

Public Function Bar() As Variant '@TestMethod
'@MemberAttribute VB_Attribute1, False
End Function
'@MemberAttribute VB_Attribute2, False";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function).Single();

                var expectedAnnotationCount = 2;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void MemberAnnotationsOnOrAbovePreviousMemberDoNotGetScopedToMember()
        {
            const string inputCode =
                @"
'@Description ""Desc""
Public Sub Foo()
End Sub _
'@TestMethod 
'SomeComment
'@Enumerator _
17 _
, _
12 _
 _
@DefaultMember _

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function).Single();

                var expectedAnnotationCount = 2;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void MemberAnnotationsOnOrAboveModuleVariableDoNotGetScopedToMember()
        {
            const string inputCode =
                @"
'@Description ""Desc""
Public foo As Long _
'@TestMethod 
'SomeComment
'@Enumerator _
17 _
, _
12 _
 _
@DefaultMember _

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function).Single();

                var expectedAnnotationCount = 2;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void ModuleAnnotationsAboveNonAnnotationLineAboveFirstMemberAreModuleAnnotations()
        {
            const string inputCode =
                @"
Public Foobar As Long
'@TestModule
'@Folder(""Test"")
'SomeComment
'@Enumerator 17, _
12 _
@DefaultMember
Public Sub Foo()
End Sub

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleDeclaration =
                    state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 2;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void NonModuleAnnotationsAboveNonAnnotationLineAboveFirstMemberAreNotModuleAnnotations()
        {
            const string inputCode =
                @"
Public Foobar As Long
'@TestMethod
'SomeComment
'@Enumerator 17, _
12 _
@DefaultMember
Public Sub Foo()
End Sub

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleDeclaration =
                    state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void ModuleAnnotationsOnOrBelowFirstMemberAreNotModuleAnnotations()
        {
            const string inputCode =
                @"
Public Foobar As Long

Public Sub Foo() '@ModuleDescription ""Desc""
End Sub

'@TestModule

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleDeclaration =
                    state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void AllModuleAnnotationsAreModuleAnnotationsIfThereIsNoBody()
        {
            const string inputCode =
                @"
'@TestModule
'@Folder(""Test"")
'SomeComment
'@Enumerator 17, _
12 _
@DefaultMember
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleDeclaration =
                    state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 2;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void VariableAnnotationsAboveVariableGetScopedToVariable()
        {
            const string inputCode =
                @"
'@Obsolete
Public foo As Long

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();

                var expectedAnnotationCount = 1;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void NonVariableAnnotationsAboveVariableDoNotGetScopedToVariable()
        {
            const string inputCode =
                @"
'@TestModule
Public foo As Long

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void LineContinuedVariableAnnotationsAboveVariableGetScopedToVariable()
        {
            const string inputCode =
                @"
'@Obsolete _

Public foo As Long

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();

                var expectedAnnotationCount = 1;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void VariableAnnotationsOnOrAbovePreviousVariableDoNotGetScopedToVariable()
        {
            const string inputCode =
                @"'@Obsolete
Private fooBar As Variant _
'@Obsolete

'SomeComment
Public foo As Long

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single(decl => decl.IdentifierName == "foo");

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = fooDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void VariableAnnotationsOnOrAboveNonWhiteSpaceStatementDoNotGetScopedToVariable()
        {
            const string inputCode =
                @"'@Obsolete
Option Explicit _
'@Obsolete
'SomeComment
Public foo As Long

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void VariableAnnotationsOnOrBelowVariableDoNotGetScopedToVariable()
        {
            const string inputCode =
                @"
'SomeComment
Public foo As Long '@Obsolete
'@Obsolete

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void VariableAnnotationsAboveFirstNonAnnotationLineAboveVariableStillGetScopedToVariable()
        {
            const string inputCode =
                @"
'@Obsolete
'SomeComment
Public foo As Long

Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var barDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();

                var expectedAnnotationCount = 1;
                var actualAnnotationCount = barDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void IdentifierAnnotationsInWhiteSpaceAboveIdentifierGetScopedToIdentifier()
        {
            const string inputCode =
                @"
Public foo As Long

Public Function Bar() As Variant

'@Ignore MissingAttribute
'Some Comment

'@TestModule

'@Ignore EmptyModule
    foo = 42
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();
                var fooReference = fooDeclaration.References.Single();

                var expectedAnnotationCount = 2;
                var actualAnnotationCount = fooReference.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void IdentifierAnnotationsOnOrBelowIdentifierDoNotGetScopedToIdentifier()
        {
            const string inputCode =
                @"
Public foo As Long

Public Function Bar() As Variant
    foo = 42 '@Ignore MissingAttribute
'@Ignore EmptyModule
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();
                var fooReference = fooDeclaration.References.Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = fooReference.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void IdentifierAnnotationsOnPreviousNonWhiteSpaceDoNotGetScopedToIdentifier()
        {
            const string inputCode =
                @"
Public foo As Long

Public Function Bar() As Variant '@Ignore MissingAttribute
    foo = 42 
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var fooDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single();
                var fooReference = fooDeclaration.References.Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = fooReference.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        //Cf. issue #5071 at https://github.com/rubberduck-vba/Rubberduck/issues/5071
        public void AnnotationArgumentIsRecognisedWithWhiteSpaceInBetween()
        {
            const string inputCode =
                @"
'@description (""Function description"")
Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function).Single();
                var annotation = declaration.Annotations.Where(pta => pta.Annotation is DescriptionAnnotation).Single();

                var expectedAnnotationArgument = "\"Function description\"";
                var actualAnnotationArgument = annotation.AnnotationArguments[0];

                Assert.AreEqual(expectedAnnotationArgument, actualAnnotationArgument);
            }
        }

        [Test]
        public void AnnotationArgumentIsRecognisedWithLineContinuationsInBetween()
        {
            const string inputCode =
                @"
'@description _
 _
 (""Function description"")
Public Function Bar() As Variant
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declaration = state.DeclarationFinder.UserDeclarations(DeclarationType.Function).Single();
                var annotation = declaration.Annotations.Where(pta => pta.Annotation is DescriptionAnnotation).Single();

                var expectedAnnotationArgument = "\"Function description\"";
                var actualAnnotationArgument = annotation.AnnotationArguments[0];

                Assert.AreEqual(expectedAnnotationArgument, actualAnnotationArgument);
            }
        }
    }
}