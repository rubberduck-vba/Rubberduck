using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class SuspiciousPredeclaredInstanceAccessInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ImplicitQualifier_NoResult()
        {
            var className = "UserForm1";
            var code = @"
Attribute VB_PredeclaredId = True
Public AnyField As Long
Private Sub Test()
    AnyField = 42
End Sub
";
            var inspectionResults = InspectionResultsForModules((className, code, ComponentType.UserForm));
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExplicitQualifierLHS_HasResult()
        {
            var className = "UserForm1";
            var code = $@"
Attribute VB_PredeclaredId = True
Public AnyField As Long
Private Sub Test()
    {className}.AnyField = 42
End Sub
";
            var inspectionResults = InspectionResultsForModules((className, code, ComponentType.UserForm));
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExplicitQualifierRHS_HasResult()
        {
            var className = "UserForm1";
            var code = $@"
Attribute VB_PredeclaredId = True
Public AnyField As Long
Private Sub Test()
    AnyField = {className}.AnyField + 42
End Sub
";
            var inspectionResults = InspectionResultsForModules((className, code, ComponentType.UserForm));
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExplicitMeQualifier_NoResult()
        {
            var className = "UserForm1";
            var code = $@"
Attribute VB_PredeclaredId = True
Public AnyField As Long
Private Sub Test()
    Me.AnyField = 42
End Sub
";
            var inspectionResults = InspectionResultsForModules((className, code, ComponentType.UserForm));
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void NoMemberAccess_NoResult()
        {
            var className = "UserForm1";
            var code = $@"
Attribute VB_PredeclaredId = True
Public AnyField As Long
Private Sub Test()
    If Me Is UserForm1 Then Exit Sub
End Sub
";
            var inspectionResults = InspectionResultsForModules((className, code, ComponentType.UserForm));
            Assert.AreEqual(0, inspectionResults.Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new SuspiciousPredeclaredInstanceAccessInspection(state);
        }
    }
}
