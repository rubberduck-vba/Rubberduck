using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class SynchronizeAttributesQuickFixTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
        public void AddsMissingPredeclaredIdAnnotation()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Attribute VB_PredeclaredId = True
Option Explicit

Sub DoSomething()
End Sub";
            const string expectedCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Attribute VB_PredeclaredId = True
'@PredeclaredId
Option Explicit

Sub DoSomething()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var result = inspector.FindIssuesAsync(state, CancellationToken.None).Result?.SingleOrDefault();
                if (result?.Context.GetType() != typeof(VBAParser.AttributeStmtContext))
                {
                    Assert.Inconclusive("Inspection failed to return a result.");
                }

                var fix = new SynchronizeAttributesQuickFix(state);
                fix.Fix(result);

                var rewriter = state.GetRewriter(result.QualifiedSelection.QualifiedName);
                var actual = rewriter.GetText();

                Assert.AreEqual(expectedCode, actual);
            }
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void AddsMissingDescriptionAnnotation()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Option Explicit

Sub DoSomething()
Attribute DoSomething.VB_Description = ""Does something""
End Sub";
            const string expectedCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Option Explicit

'@Description(""Does something"")
Sub DoSomething()
Attribute DoSomething.VB_Description = ""Does something""
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var result = inspector.FindIssuesAsync(state, CancellationToken.None).Result?.SingleOrDefault();
                if (result?.Context.GetType() != typeof(VBAParser.AttributeStmtContext))
                {
                    Assert.Inconclusive("Inspection failed to return a result.");
                }

                var fix = new SynchronizeAttributesQuickFix(state);
                fix.Fix(result);

                var rewriter = state.GetRewriter(result.QualifiedSelection.QualifiedName);
                var actual = rewriter.GetText();

                Assert.AreEqual(expectedCode, actual);
            }
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void AddsDefaultMemberAnnotation()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Option Explicit

Sub DoSomething()
Attribute DoSomething.VB_UserMemId = 0
End Sub";
            const string expectedCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Option Explicit

'@DefaultMember
Sub DoSomething()
Attribute DoSomething.VB_UserMemId = 0
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var result = inspector.FindIssuesAsync(state, CancellationToken.None).Result?.SingleOrDefault();
                if (result?.Context.GetType() != typeof(VBAParser.AttributeStmtContext))
                {
                    Assert.Inconclusive("Inspection failed to return a result.");
                }

                var fix = new SynchronizeAttributesQuickFix(state);
                fix.Fix(result);

                var rewriter = state.GetRewriter(result.QualifiedSelection.QualifiedName);
                var actual = rewriter.GetText();

                Assert.AreEqual(expectedCode, actual);
            }
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void AddsEnumeratorMemberAnnotation()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Option Explicit

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
End Property";
            const string expectedCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Option Explicit

'@Enumerator
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
End Property";
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var result = inspector.FindIssuesAsync(state, CancellationToken.None).Result?.SingleOrDefault();
                if (result?.Context.GetType() != typeof(VBAParser.AttributeStmtContext))
                {
                    Assert.Inconclusive("Inspection failed to return a result.");
                }

                var fix = new SynchronizeAttributesQuickFix(state);
                fix.Fix(result);

                var rewriter = state.GetRewriter(result.QualifiedSelection.QualifiedName);
                var actual = rewriter.GetText();

                Assert.AreEqual(expectedCode, actual);
            }
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void AddsMissingPredeclaredIdAttribute()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@PredeclaredId
";
            const string expectedCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@PredeclaredId
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAttributeInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var result = inspector.FindIssuesAsync(state, CancellationToken.None).Result?.SingleOrDefault();
                if (result?.Context.GetType() != typeof(VBAParser.AnnotationContext))
                {
                    Assert.Inconclusive("Inspection failed to return a result.");
                }

                var fix = new SynchronizeAttributesQuickFix(state);
                fix.Fix(result);

                var rewriter = state.GetAttributeRewriter(result.QualifiedSelection.QualifiedName);
                var actual = rewriter.GetText();

                Assert.AreEqual(expectedCode, actual);
            }
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void AddsMissingExposedAttribute()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Exposed
";
            const string expectedCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'@Exposed
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAttributeInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var result = inspector.FindIssuesAsync(state, CancellationToken.None).Result?.SingleOrDefault();
                if (result?.Context.GetType() != typeof(VBAParser.AnnotationContext))
                {
                    Assert.Inconclusive("Inspection failed to return a result.");
                }

                var fix = new SynchronizeAttributesQuickFix(state);
                fix.Fix(result);

                var rewriter = state.GetAttributeRewriter(result.QualifiedSelection.QualifiedName);
                var actual = rewriter.GetText();

                Assert.AreEqual(expectedCode, actual);
            }
        }
    }
}