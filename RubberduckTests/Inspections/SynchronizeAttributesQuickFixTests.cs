using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class SynchronizeAttributesQuickFixTests
    {
        [Ignore] // todo: implement
        [TestMethod]
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
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out component);

            var state = MockParser.CreateAndParse(vbe.Object);
            var inspection = new MissingAnnotationInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var result = inspector.FindIssuesAsync(state, CancellationToken.None).Result?.SingleOrDefault();
            if(result?.Context.GetType() != typeof(VBAParser.AttributeStmtContext))
            {
                Assert.Inconclusive("Inspection failed to return a result.");
            }

            var fix = new SynchronizeAttributesQuickFix(state);
            fix.Fix(result);

            var rewriter = state.GetAttributeRewriter(result.QualifiedSelection.QualifiedName);
            var actual = rewriter.GetText();

            Assert.AreEqual(expectedCode, actual);

        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out component);

            var state = MockParser.CreateAndParse(vbe.Object);
            var inspection = new MissingAttributeInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var result = inspector.FindIssuesAsync(state, CancellationToken.None).Result?.SingleOrDefault();
            if(result?.Context.GetType() != typeof(VBAParser.AnnotationContext))
            {
                Assert.Inconclusive("Inspection failed to return a result.");
            }

            var fix = new SynchronizeAttributesQuickFix(state);
            fix.Fix(result);

            var rewriter = state.GetAttributeRewriter(result.QualifiedSelection.QualifiedName);
            var actual = rewriter.GetText();

            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out component);

            var state = MockParser.CreateAndParse(vbe.Object);
            var inspection = new MissingAttributeInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var result = inspector.FindIssuesAsync(state, CancellationToken.None).Result?.SingleOrDefault();
            if(result?.Context.GetType() != typeof(VBAParser.AnnotationContext))
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