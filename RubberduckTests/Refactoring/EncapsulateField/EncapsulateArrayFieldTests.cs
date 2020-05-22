using System;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateArrayFieldTests : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase("Private", "mArray(5) As String", "mArray(5) As String")]
        [TestCase("Public", "mArray(5) As String", "mArray(5) As String")]
        [TestCase("Private", "mArray(5,2,3) As String", "mArray(5,2,3) As String")]
        [TestCase("Public", "mArray(5,2,3) As String", "mArray(5,2,3) As String")]
        [TestCase("Private", "mArray(1 to 10) As String", "mArray(1 to 10) As String")]
        [TestCase("Public", "mArray(1 to 10) As String", "mArray(1 to 10) As String")]
        [TestCase("Private", "mArray() As String", "mArray() As String")]
        [TestCase("Public", "mArray() As String", "mArray() As String")]
        [TestCase("Private", "mArray(5)", "mArray(5) As Variant")]
        [TestCase("Public", "mArray(5)", "mArray(5) As Variant")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateArray(string visibility, string arrayDeclaration, string expectedArrayDeclaration)
        {
            string inputCode =
                $@"Option Explicit

{visibility} {arrayDeclaration}";

            var selection = new Selection(3, 8, 3, 11);

            string expectedCode =
                $@"Option Explicit

Private {expectedArrayDeclaration}

Public Property Get MyArray() As Variant
    MyArray = mArray
End Property
";
            var userInput = new UserInputDataObject()
                .UserSelectsField("mArray", "MyArray");

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            var expectedLines = expectedCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            Assert.AreEqual(expectedCode.Trim(), actualCode);
        }

        [TestCase("5")]
        [TestCase("5,2,3")]
        [TestCase("1 to 100")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateArray_DeclaredInList(string dimensions)
        {
            string inputCode =
                $@"Option Explicit

Public mArray({dimensions}) As String, anotherVar As Long, andOneMore As Variant";

            var selection = new Selection(3, 8, 3, 11);

            string expectedCode =
                $@"Option Explicit

Public anotherVar As Long, andOneMore As Variant
Private mArray({dimensions}) As String

Public Property Get MyArray() As Variant
    MyArray = mArray
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("mArray", "MyArray");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            StringAssert.Contains("Public anotherVar As Long, andOneMore As Variant", actualCode);
            StringAssert.Contains($"Private mArray({dimensions}) As String", actualCode);
            StringAssert.Contains("Get MyArray() As Variant", actualCode);
            StringAssert.Contains("MyArray = mArray", actualCode);
            StringAssert.DoesNotContain("Let MyArray", actualCode);
            StringAssert.DoesNotContain("Set MyArray", actualCode);
        }

        [TestCase("mArr|ay(5) As String, mNextVar As Long", "Private mArray(5) As String")]
        [TestCase("mNextVar As Long, mArr|ay(5) As String", "Private mArray(5) As String")]
        [TestCase("mArr|ay(5), mNextVar As Long", "Private mArray(5) As Variant")]
        [TestCase("mNextVar As Long, mAr|ray(5)", "Private mArray(5) As Variant")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateArray_newFieldNameForFieldInList(string declarationList, string expectedDeclaration)
        {
            string inputCode =
                $@"Option Explicit

Public {declarationList}";

            string expectedCode =
                $@"Option Explicit

Public mNextVar As Long

{expectedDeclaration}

Public Property Get MyArray() As Variant
    MyArray = mArray
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("mArray", "MyArray");
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode.Trim(), actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void RedimsBackingVariable()
        {
            string inputCode =
                $@"Option Explicit

Public myA|rray() As Integer

Private Sub InitializeArray(size As Long)
    Redim myArray(size)
    Dim idx As Long
    For idx = 1 To size  
        myArray(idx) = idx 
    Next idx
End Sub
";

            var presenterAction = Support.UserAcceptsDefaults();
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Public Property Get MyArray() As Variant", actualCode);
            StringAssert.DoesNotContain("Public Property Let MyArray(", actualCode);
            StringAssert.Contains("Redim myArray_1(size)", actualCode);
            StringAssert.Contains("myArray_1(idx) = idx", actualCode);
        }


        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void RedimsBackingVariableAsUDT()
        {
            string inputCode =
                $@"Option Explicit

Public myA|rray() As Integer

Private Sub InitializeArray(size As Long)
    Redim myArray(size)
    Dim idx As Long
    For idx = 1 To size  
        myArray(idx) = idx 
    Next idx
End Sub
";

            var presenterAction = Support.UserAcceptsDefaults(convertFieldToUDTMember: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Public Property Get MyArray() As Variant", actualCode);
            StringAssert.DoesNotContain("Public Property Let MyArray(", actualCode);
            StringAssert.Contains("Redim this.MyArray(size)", actualCode);
            StringAssert.Contains("this.MyArray(idx) = idx", actualCode);
        }

        [TestCase(false)]
        [TestCase(true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void RedimsBackingVariableExternally(bool convertField)
        {
            var fieldUT = "myArray";
            string inputCode =
                $@"Option Explicit

Public myArray() As Long
";
            var redimCode =
                $@"Option Explicit

Private Sub InitializeArray(size As Long)
    Redim myArray(size)
    Dim idx As Long
    For idx = 1 To size  
        myArray(idx) = idx 
    Next idx
End Sub
";
            var presenterAction = Support.UserAcceptsDefaults(convertFieldToUDTMember: convertField);

            var vbe = MockVbeBuilder.BuildFromStdModules(("SourceModule", inputCode), ("ClientModule", redimCode));
            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(vbe.Object, fieldUT, DeclarationType.Variable, presenterAction);

            var field = model[fieldUT];

            field.TryValidateEncapsulationAttributes(out var errorMessage);

            var expectedError = string.Format(RubberduckUI.EncapsulateField_ArrayHasExternalRedimFormat, field.IdentifierName);

            StringAssert.AreEqualIgnoringCase(expectedError, errorMessage);
        }

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            RefactoringUserInteraction<IEncapsulateFieldPresenter, EncapsulateFieldModel> userInteraction,
            ISelectionService selectionService)
        {
            return Support.SupportTestRefactoring(rewritingManager, state, userInteraction, selectionService);
        }
    }
}
