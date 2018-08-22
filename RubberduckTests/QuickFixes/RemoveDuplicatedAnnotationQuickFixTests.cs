using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveDuplicatedAnnotationQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void RemoveDuplicatedAnnotation_QuickFixWorks_RemoveDuplicate()
        {
            const string inputCode = @"
'@Obsolete
'@Obsolete
Public Sub Foo
End Sub";

            const string expectedCode = @"
'
'@Obsolete
Public Sub Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DuplicatedAnnotationInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new RemoveDuplicatedAnnotationQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RemoveDuplicatedAnnotation_QuickFixWorks_RemoveMultipleDuplicates()
        {
            const string inputCode = @"
'@Obsolete
'@Obsolete
'@Obsolete
Public Sub Foo
End Sub";

            const string expectedCode = @"
'
'
'@Obsolete
Public Sub Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DuplicatedAnnotationInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new RemoveDuplicatedAnnotationQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RemoveDuplicatedAnnotation_QuickFixWorks_RemoveDuplicateWithComment()
        {
            const string inputCode = @"
'@Obsolete: Foo
'@Obsolete
Public Sub Foo
End Sub";

            const string expectedCode = @"
': Foo
'@Obsolete
Public Sub Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DuplicatedAnnotationInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new RemoveDuplicatedAnnotationQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RemoveDuplicatedAnnotation_QuickFixWorks_RemoveDuplicateFromSameAnnotationList()
        {
            const string inputCode = @"
'@Obsolete @Obsolete
Public Sub Foo
End Sub";

            const string expectedCode = @"
'@Obsolete
Public Sub Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DuplicatedAnnotationInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new RemoveDuplicatedAnnotationQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RemoveDuplicatedAnnotation_QuickFixWorks_RemoveMultipleDuplicatesFromSameAnnotationList()
        {
            const string inputCode = @"
'@Obsolete @Obsolete @Obsolete
Public Sub Foo
End Sub";

            const string expectedCode = @"
'@Obsolete
Public Sub Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DuplicatedAnnotationInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new RemoveDuplicatedAnnotationQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RemoveDuplicatedAnnotation_QuickFixWorks_RemoveDuplicateFromOtherAnnotationList()
        {
            const string inputCode = @"
'@Obsolete @NoIndent
'@Obsolete
Public Sub Foo
End Sub";

            const string expectedCode = @"
'@NoIndent
'@Obsolete
Public Sub Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DuplicatedAnnotationInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new RemoveDuplicatedAnnotationQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RemoveDuplicatedAnnotation_QuickFixWorks_RemoveMultipleDuplicatesFromOtherAnnotationList()
        {
            const string inputCode = @"
'@Obsolete @NoIndent @Obsolete
'@Obsolete
Public Sub Foo
End Sub";

            const string expectedCode = @"
'@NoIndent
'@Obsolete
Public Sub Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DuplicatedAnnotationInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new RemoveDuplicatedAnnotationQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }
    }
}
