using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using RubberduckTests.Mocks;

namespace RubberduckTests.Annotations
{
    [TestFixture]
    public class AnnotationResolutionTests
    {
        [Test]
        public void AnnotationsAboveMemberGetScopedToMember_NotFirstMember()
        {
            const string inputCode =
                @"
Public Sub Foo()
End Sub
'@TestMethod
'@Enumerator 17, 12 @DefaultMember
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
        public void AnnotationsAboveMemberGetScopedToMember_FirstMember()
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
        public void LineContinuedAnnotationsAboveMemberGetScopedToMember_NotFirstMember()
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
        public void LineContinuedAnnotationsAboveMemberGetScopedToMember_FirstMember()
        {
            const string inputCode =
                @"
'@TestMethod _

'@Enumerator _
17 _
, _
12 _
 _
@DefaultMember _

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
        public void AnnotationsRightAboveFirstMemberAreNotModuleAnnotations_WithDeclarationOnTop()
        {
            const string inputCode =
                @"
Public Foobar As Long

'@TestMethod
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
                var moduleDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void AnnotationsRightAboveFirstMemberAreNotModuleAnnotations_WithoutDeclarationOnTop()
        {
            const string inputCode =
                @"'@TestMethod
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
                var moduleDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 0;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void AnnotationsAboveNonAnnotationLineAboveFirstMemberAreModuleAnnotations_WithDeclarationOnTop()
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
                var moduleDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 2;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void AnnotationsAboveNonAnnotationLineAboveFirstMemberAreModuleAnnotations_WithoutDeclarationOnTop()
        {
            const string inputCode =
                @"
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
                var moduleDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 2;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }

        [Test]
        public void AllAnnotationsAreModuleAnnotationsIfThereIsNoBody()
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
                var moduleDeclaration = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single();

                var expectedAnnotationCount = 4;
                var actualAnnotationCount = moduleDeclaration.Annotations.Count();

                Assert.AreEqual(expectedAnnotationCount, actualAnnotationCount);
            }
        }
    }
}