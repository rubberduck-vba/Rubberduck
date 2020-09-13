using NUnit.Framework;
using Rubberduck.Refactorings.EncapsulateField;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var encapsulateTarget = state.AllUserDeclarations.Single(d => d.IdentifierName.Equals("mVehicle"));
                var objectStateUDTTarget = state.AllUserDeclarations.Single(d => d.IdentifierName.Equals("this"));

                var resolver = new EncapsulateFieldTestComponentResolver(state, null);

                var encapsulateFieldCandidateFactory = resolver.Resolve<IEncapsulateFieldCandidateFactory>();
                var objStateFactory = resolver.Resolve<IObjectStateUserDefinedTypeFactory>();

                var objStateCandidate = encapsulateFieldCandidateFactory.Create(objectStateUDTTarget);
                var objStateUDT = objStateFactory.Create(objStateCandidate as IUserDefinedTypeCandidate);

                var candidate = new EncapsulateFieldAsUDTMemberCandidate(encapsulateFieldCandidateFactory.Create(encapsulateTarget), objStateUDT)
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var encapsulateTarget = state.AllUserDeclarations.Single(d => d.IdentifierName.Equals("mTest"));
                var objectStateUDTTarget = state.AllUserDeclarations.Single(d => d.IdentifierName.Equals("this"));

                var resolver = new EncapsulateFieldTestComponentResolver(state, null);

                var encapsulateFieldCandidateFactory = resolver.Resolve<IEncapsulateFieldCandidateFactory>();
                var objStateFactory = resolver.Resolve<IObjectStateUserDefinedTypeFactory>();

                var objStateCandidate = encapsulateFieldCandidateFactory.Create(objectStateUDTTarget);
                var objStateUDT = objStateFactory.Create(objStateCandidate as IUserDefinedTypeCandidate);

                var candidate = new EncapsulateFieldAsUDTMemberCandidate(encapsulateFieldCandidateFactory.Create(encapsulateTarget), objStateUDT);

                var generator = new PropertyAttributeSetsGenerator();
                var propAttributeSets = generator.GeneratePropertyAttributeSets(candidate);
                StringAssert.Contains("this.Test.Number4Type.Number3Type.Number2Type.Number1Type.DeeplyNested", propAttributeSets.First().BackingField);
            }
        }
    }
}
