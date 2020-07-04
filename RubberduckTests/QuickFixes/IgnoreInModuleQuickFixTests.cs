using System.Collections.Generic;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class IgnoreInModuleQuickFixTests :  QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void NoIgnoreModule_AddsNewOne()
        {
            var inputCode =
                @"
Public Sub DoSomething()
    Dim value As Long
    Dim bar As Long
    value = 42
    bar = 23
    Debug.Print 42
End Sub";

            var expectedCode =
                @"'@IgnoreModule VariableNotUsed

Public Sub DoSomething()
    Dim value As Long
    Dim bar As Long
    value = 42
    bar = 23
    Debug.Print 42
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IgnoreModuleAlreadyThere_InspectionNotIgnoredYet_AddsNewArgument()
        {
            var inputCode =
                @"'@IgnoreModule AssignmentNotUsed
Public Sub DoSomething()
    Dim value As Long
    Dim bar As Long
    value = 42
    bar = 23
    Debug.Print 42
End Sub";

            var expectedCode =
                @"'@IgnoreModule VariableNotUsed, AssignmentNotUsed
Public Sub DoSomething()
    Dim value As Long
    Dim bar As Long
    value = 42
    bar = 23
    Debug.Print 42
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IgnoreModuleMultiple_AddsOneAnnotation_NoIgnoreModuleYet()
        {
            var inputCode =
                @"
Public Sub DoSomething()
    Dim value As Long
    Dim bar As Long
    value = 42
    bar = 23
    Debug.Print 42
End Sub";

            var expectedCode =
                @"'@IgnoreModule VariableNotUsed

Public Sub DoSomething()
    Dim value As Long
    Dim bar As Long
    value = 42
    bar = 23
    Debug.Print 42
End Sub";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new VariableNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IgnoreModuleMultiple_AddsOneAnnotation_IgnoreModuleAlreadyThere()
        {
            var inputCode =
                @"'@IgnoreModule AssignmentNotUsed
Public Sub DoSomething()
    Dim value As Long
    Dim bar As Long
    value = 42
    bar = 23
    Debug.Print 42
End Sub";

            var expectedCode =
                @"'@IgnoreModule VariableNotUsed, AssignmentNotUsed
Public Sub DoSomething()
    Dim value As Long
    Dim bar As Long
    value = 42
    bar = 23
    Debug.Print 42
End Sub";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new VariableNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            var annotationUpdater = new AnnotationUpdater(state);
            var inspections = new List<IInspection> {new VariableNotUsedInspection(state)};
            return new IgnoreInModuleQuickFix(annotationUpdater, state, inspections);
        }
    }
}