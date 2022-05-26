using NUnit.Framework;
using Rubberduck.Refactorings.EncapsulateField;
using System.Linq;
using RubberduckTests.Mocks;
using Rubberduck.Refactorings;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class PropertyAttributeSetsGeneratorTests
    {
        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(PropertyAttributeSetsGenerator))]
        public void EncapsulateFieldCandidate_PrivateUDTField()
        {
            var inputCode =
$@"
Option Explicit

Private Type TVehicle
    Wheels As Integer
End Type

Private Type TObjState
    FirstValue As String
End Type

Private this As TObjState

Private mVehicle As TVehicle
";

            
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var encapsulateTarget = state.AllUserDeclarations.Single(d => d.IdentifierName.Equals("mVehicle"));
                var objectStateUDTTarget = state.AllUserDeclarations.Single(d => d.IdentifierName.Equals("this"));

                var encapsulateFieldCandidateFactory = EncapsulateFieldTestSupport.GetResolver(state)
                    .Resolve<IEncapsulateFieldCandidateFactory>();

                var objStateCandidate = encapsulateFieldCandidateFactory.CreateFieldCandidate(objectStateUDTTarget);
                var objStateUDT = encapsulateFieldCandidateFactory.CreateObjectStateField(objStateCandidate as IUserDefinedTypeCandidate);

                var candidate = new EncapsulateFieldAsUDTMemberCandidate(encapsulateFieldCandidateFactory.CreateFieldCandidate(encapsulateTarget), objStateUDT)
                {
                    PropertyIdentifier = "MyType"
                };

                var generator = new PropertyAttributeSetsGenerator();
                var propAttributeSets = generator.GeneratePropertyAttributeSets(candidate);
                StringAssert.Contains("this.MyType.Wheels", propAttributeSets.First().BackingField);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(PropertyAttributeSetsGenerator))]
        public void EncapsulateFieldCandidate_DeeplyNestedUDTs()
        {
            var inputCode =
$@"
Option Explicit

Private Type FirstType
    DeeplyNested As Long
End Type

Private Type SecondType
    Number1Type As FirstType
End Type

Private Type ThirdType
    Number2Type As SecondType
End Type

Private Type FourthType
    Number3Type As ThirdType
End Type

Private Type FifthType
    Number4Type As FourthType
End Type

Private Type ExistingType
    ExistingValue As String
End Type

Public mTest As FifthType

Private this As ExistingType

";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var encapsulateTarget = state.AllUserDeclarations.Single(d => d.IdentifierName.Equals("mTest"));
                var objectStateUDTTarget = state.AllUserDeclarations.Single(d => d.IdentifierName.Equals("this"));

                var encapsulateFieldCandidateFactory = EncapsulateFieldTestSupport.GetResolver(state)
                    .Resolve<IEncapsulateFieldCandidateFactory>();

                var objStateCandidate = encapsulateFieldCandidateFactory.CreateFieldCandidate(objectStateUDTTarget);
                var objStateUDT = encapsulateFieldCandidateFactory.CreateObjectStateField(objStateCandidate as IUserDefinedTypeCandidate);

                var candidate = new EncapsulateFieldAsUDTMemberCandidate(encapsulateFieldCandidateFactory.CreateFieldCandidate(encapsulateTarget), objStateUDT);

                var generator = new PropertyAttributeSetsGenerator();
                var propAttributeSets = generator.GeneratePropertyAttributeSets(candidate);
                StringAssert.Contains("this.Test.Number4Type.Number3Type.Number2Type.Number1Type.DeeplyNested", propAttributeSets.First().BackingField);
            }
        }
    }
}
