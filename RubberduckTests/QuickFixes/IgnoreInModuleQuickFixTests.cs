using System.Collections.Generic;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

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

        [Test]
        [Category("QuickFixes")]
        public void IgnoreModuleMultiple_IdentifierInspection_DeclarationInOtherModule_AnnotationInReferenceModule()
        {
            var classCode =
                @"
'@Obsolete(""no longer use this"")
Public Sub Foo()
End Sub"; 
            
            var moduleCode =
                 @"Public Bar As Class1";

            var inputCode =
                @"
Public Sub DoSomething()
    Module1.Bar.Foo
End Sub";

            var expectedCode =
                @"'@IgnoreModule ObsoleteMemberUsage

Public Sub DoSomething()
    Module1.Bar.Foo
End Sub";

            var vbe = new MockVbeBuilder().ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, classCode)
                .AddComponent("Module1", ComponentType.StandardModule, moduleCode)
                .AddComponent("TestModule", ComponentType.StandardModule, inputCode)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;
            var actualCode = ApplyQuickFixToFirstInspectionResult(vbe, "TestModule", state => new ObsoleteMemberUsageInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IgnoreMcoduleMultiple_IdentifierInspection_DeclarationInOtherModule_LeavesDeclarationModuleAsIs()
        {
            var classCode =
                @"
'@Obsolete(""no longer use this"")
Public Sub Foo()
End Sub";

            var moduleCode =
                 @"Public Bar As Class1";

            var inputCode =
                @"
Public Sub DoSomething()
    Module1.Bar.Foo
End Sub";

            var vbe = new MockVbeBuilder().ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, classCode)
                .AddComponent("Module1", ComponentType.StandardModule, moduleCode)
                .AddComponent("TestModule", ComponentType.StandardModule, inputCode)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;
            var actualCode = ApplyQuickFixToFirstInspectionResult(vbe, "Class1", state => new ObsoleteMemberUsageInspection(state));
            Assert.AreEqual(classCode, actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            var annotationUpdater = new AnnotationUpdater(state);
            var inspections = new List<IInspection> {new VariableNotUsedInspection(state)};
            return new IgnoreInModuleQuickFix(annotationUpdater, state, inspections);
        }
    }
}