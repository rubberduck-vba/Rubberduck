using System;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveDuplicatedAnnotationQuickFixTests : QuickFixTestBase
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
'@Obsolete
Public Sub Foo
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new DuplicatedAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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
'@Obsolete
Public Sub Foo
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new DuplicatedAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RemoveDuplicatedAnnotation_QuickFixWorks_RemoveDuplicateWithComment()
        {
            const string inputCode = @"
'@Obsolete
'@Obsolete: Foo
Public Sub Foo
End Sub";

            const string expectedCode = @"
'@Obsolete
' Foo
Public Sub Foo
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new DuplicatedAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new DuplicatedAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new DuplicatedAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RemoveDuplicatedAnnotation_QuickFixWorks_RemoveDuplicateFromOtherAnnotationList()
        {
            const string inputCode = @"
'@Obsolete
'@Obsolete @NoIndent
Public Sub Foo
End Sub";

            const string expectedCode = @"
'@Obsolete
'@NoIndent
Public Sub Foo
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new DuplicatedAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RemoveDuplicatedAnnotation_QuickFixWorks_RemoveMultipleDuplicatesFromOtherAnnotationList()
        {
            const string inputCode = @"
'@Obsolete
'@Obsolete @NoIndent @Obsolete
Public Sub Foo
End Sub";

            const string expectedCode = @"
'@Obsolete
'@NoIndent 
Public Sub Foo
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new DuplicatedAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RemoveDuplicatedAnnotation_QuickFixWorks_RemoveDuplicatesWithoutWhitespaceFromAnnotationList()
        {
            const string inputCode = @"
'@Obsolete@Obsolete
Public Sub Foo
End Sub";

            const string expectedCode = @"
'@Obsolete
Public Sub Foo
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new DuplicatedAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RemoveDuplicatedAnnotation_QuickFixWorks_RemoveDuplicatesOfOnlyOneAnnotation()
        {
            const string inputCode = @"
'@Obsolete
'@Obsolete
'@TestMethod
'@TestMethod
Public Sub Foo
End Sub";

            const string expectedCode = @"
'@Obsolete
'@TestMethod
'@TestMethod
Public Sub Foo
End Sub";
            Func<IInspectionResult, bool> conditionToFix = result => result is IWithInspectionResultProperties<IAnnotation> resultProperties 
                                                                     && resultProperties.Properties is ObsoleteAnnotation;
            var actualCode = ApplyQuickFixToFirstInspectionResultSatisfyingPredicate(inputCode, state => new DuplicatedAnnotationInspection(state), conditionToFix);
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveDuplicatedAnnotationQuickFix(new AnnotationUpdater(state));
        }
    }
}
