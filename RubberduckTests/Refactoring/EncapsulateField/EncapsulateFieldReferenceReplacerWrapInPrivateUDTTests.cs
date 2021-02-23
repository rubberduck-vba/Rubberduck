using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateFieldReferenceReplacerWrapInPrivateUDTTests
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void NestedUDTMember_WrapInPrivateUDT(bool isReadOnly)
        {
            var target = "mTypesField";

            var testTargetTuple = (target, "TypesField", isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Option Explicit

Private Type PType1
    FirstValType1 As Long
    SecondValType1 As String
End Type


Private Type PType2
    FirstValType2 As Long
    SecondValType2 As String
    Third As PType1
End Type

Private mTypesField As PType2

Private Sub Class_Initialize()
    mTypesField.Third.SecondValType1 = ""Wah""
End Sub

Private Sub TestSub2()
    TestSub3 mTypesField.Third.SecondValType1
End Sub

Private Sub TestSub3(ByVal arg As String)
End Sub
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule);

            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);

            var expectedAssignment = isReadOnly ? "this.TypesField.Third.SecondValType1 = \"Wah\"" : "SecondValType1 = \"Wah\"";
            StringAssert.Contains(expectedAssignment, refactoredCode[testModuleName]);
            StringAssert.Contains("TestSub3 SecondValType1", refactoredCode[testModuleName]);
        }
        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void EncapsulatePublicFields_NestedPathForPrivateUDTFieldReadonlyFlag(bool isReadOnly)
        {
            var target = "mVehicle";

            var testTargetTuple = (target, "Vehicle", isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Option Explicit

Private Type TVehicle
    Seats As Integer
    Wheels As Integer
End Type

Private mVehicle As TVehicle

Private Sub Class_Initialize()
    mVehicle.Wheels = 4
    mVehicle.Seats = 2
End Sub
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule);
            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);
            var result = refactoredCode[MockVbeBuilder.TestModuleName];

            var expectedAssignmentExpressionWheels = isReadOnly ? " this.Vehicle.Wheels = 4" : " Wheels = 4";
            StringAssert.Contains(expectedAssignmentExpressionWheels, result);
            var expectedAssignmentExpressionSeats = isReadOnly ? " this.Vehicle.Seats = 2" : " Seats = 2";
            StringAssert.Contains(expectedAssignmentExpressionSeats, result);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void EncapsulatePublicFields_NestedPathForFieldReadonlyFlag(bool isReadOnly)
        {
            var target = "mWheels";

            var testTargetTuple = (target, "Wheels", isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Option Explicit

Public mWheels As Long

Private Sub Class_Initialize()
    mWheels = 7
End Sub
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule);
            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);
            var result = refactoredCode[MockVbeBuilder.TestModuleName];

            var expectedAssignmentExpression = isReadOnly ? " this.Wheels = 7" : "  Wheels = 7";
            StringAssert.Contains(expectedAssignmentExpression, result);
        }

        private IDictionary<string, string> ReplaceReferences(IVBE vbe, (string fieldID, string fieldProperty, bool readOnly) target, params (string fieldID, string fieldProperty, bool readOnly)[] fieldIDPairs)
            => ReplaceReferences(vbe, target, fieldIDPairs.ToList());
        
        private IDictionary<string, string> ReplaceReferences(IVBE vbe, (string fieldID, string fieldProperty, bool readOnly) target, IEnumerable<(string fieldID, string fieldProperty, bool readOnly)> fieldIDPairs)
        {
            var refactoredCode = new Dictionary<string, string>();
            (var state, var rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                var resolver = Support.SetupResolver(state, rewritingManager);
                var encapsulateFieldFactory = resolver.Resolve<IEncapsulateFieldCandidateFactory>();
                var sutFactory = resolver.Resolve<IEncapsulateFieldReferenceReplacerFactory>();

                var fieldDeclaration = state.DeclarationFinder.MatchName(target.fieldID).Single();
                var defaultObjectStateUDT = encapsulateFieldFactory.CreateDefaultObjectStateField(fieldDeclaration.QualifiedModuleName);
                var fieldCandidate = encapsulateFieldFactory.CreateFieldCandidate(fieldDeclaration);

                var udtfieldCandidate = encapsulateFieldFactory.CreateUDTMemberCandidate(fieldCandidate, defaultObjectStateUDT);
                udtfieldCandidate.PropertyIdentifier = target.fieldProperty;
                udtfieldCandidate.IsReadOnly = target.readOnly;
                udtfieldCandidate.EncapsulateFlag = true;

                var sut = sutFactory.Create();
                sut.ReplaceReferences(new IEncapsulateFieldCandidate[] { udtfieldCandidate }, rewriteSession);


                if (rewriteSession.TryRewrite())
                {
                    refactoredCode = vbe.ActiveVBProject.VBComponents
                        .ToDictionary(component => component.Name, component => component.CodeModule.Content());
                }
            }

            return refactoredCode;
        }
    }
}
