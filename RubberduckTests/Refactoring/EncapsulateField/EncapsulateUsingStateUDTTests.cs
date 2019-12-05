using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateUsingStateUDTTests : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_UserAcceptsDefaults_ConflictWithStateUDT(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} my|Bar As TBar

Private this As Long";


            var presenterAction = Support.UserAcceptsDefaults(asUDT: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private this As Long", actualCode);
            StringAssert.Contains("Private this1 As This_Type", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_OnlyEncapsulateUDTMembers()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private my|Bar As TBar";


            var userInput = new UserInputDataObject("myBar");
            userInput["myBar"].EncapsulateFlag = false;
            userInput.AddUDTMemberNameFlagPairs(("myBar", "First", true));
            userInput.AddUDTMemberNameFlagPairs(("myBar", "Second", true));
            userInput.EncapsulateAsUDT = true;

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("this.MyBar.First = value", actualCode);
            StringAssert.Contains($"First = this.MyBar.First", actualCode);
            StringAssert.Contains("this.MyBar.Second = value", actualCode);
            StringAssert.Contains($"Second = this.MyBar.Second", actualCode);
            StringAssert.Contains($"MyBar As TBar", actualCode);
            StringAssert.DoesNotContain($"myBar As TBar", actualCode);
        }

        [TestCase(false, "this.MyBar.First = newValue")]
        [TestCase(true, "  First = newValue")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_UDTFieldReferences(bool encapsulateFirst, string expected)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private my|Bar As TBar

Public Sub Foo(newValue As String)
    myBar.First = newValue
End Sub";


            var userInput = new UserInputDataObject("myBar");
            userInput["myBar"].EncapsulateFlag = false;
            userInput.AddUDTMemberNameFlagPairs(("myBar", "First", encapsulateFirst));
            userInput.AddUDTMemberNameFlagPairs(("myBar", "Second", true));
            userInput.EncapsulateAsUDT = true;

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void MultipleFields()
        {
            string inputCode =
$@"
Public fo|o As Long
Public bar As String
Public foobar As Byte
";

            var userInput = new UserInputDataObject("foo", encapsulationFlag: true);
            userInput.AddAttributeSet("bar", encapsulationFlag: true);
            userInput.AddAttributeSet("foobar", encapsulationFlag: true);
            userInput.EncapsulateAsUDT = true;

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private this As This_Type", actualCode);
            StringAssert.Contains("Private Type This_Type", actualCode);
            StringAssert.Contains("Foo As Long", actualCode);
            StringAssert.Contains("Bar As String", actualCode);
            StringAssert.Contains("Foobar As Byte", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_MultipleFieldsWithUDT()
        {
            string inputCode =
$@"

Private Type TBar
    First As Long
    Second As String
End Type

Public fo|o As Long
Public myBar As TBar
";

            var userInput = new UserInputDataObject("foo", encapsulationFlag: true);
            userInput.AddAttributeSet("myBar", encapsulationFlag: true);
            //userInput.AddUDTMemberNameFlagPairs(("myBar", "First", true), ("myBar", "Second", true));
            userInput.EncapsulateAsUDT = true;

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("this.MyBar.First = value", actualCode);
            StringAssert.Contains("First = this.MyBar.First", actualCode);
            StringAssert.Contains("this.MyBar.Second = value", actualCode);
            StringAssert.Contains("Second = this.MyBar.Second", actualCode);
            var index = actualCode.IndexOf("Get Second", StringComparison.InvariantCultureIgnoreCase);
            var indexLast = actualCode.LastIndexOf("Get Second", StringComparison.InvariantCultureIgnoreCase);
            Assert.AreEqual(index, indexLast);
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            return Support.SupportTestRefactoring(rewritingManager, state, factory, selectionService);
        }
    }
}
