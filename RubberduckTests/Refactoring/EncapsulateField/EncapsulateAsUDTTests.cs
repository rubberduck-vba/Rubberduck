using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateAsUDTTests
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_MultipleFields_ToNewUDT()
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
        public void UserDefinedType_MultipleFieldsWithUDT_ToNewUDT()
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
            userInput.AddUDTMemberNameFlagPairs(("myBar", "First", true), ("myBar", "Second", true));
            userInput.EncapsulateAsUDT = true;

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("this.MyBar.First = value", actualCode);
            StringAssert.Contains("First = this.MyBar.First", actualCode);
            StringAssert.Contains("this.MyBar.Second = value", actualCode);
            StringAssert.Contains("Second = this.MyBar.Second", actualCode);
        }
    }
}
