using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.UI.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Linq;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveableMemberViewModelTests
    {
        [TestCase("Fizz", "Public Property Let\\Get Fizz() As Long")]
        [TestCase("Bizz", "Public Property Set\\Get Bizz() As Collection")]
        [TestCase("Dizz", "Public Property Let\\Set\\Get Dizz() As Variant")]
        [TestCase("WriteOnlyValue", "Public Property Let WriteOnlyValue(ByVal arg As Long)")]
        [TestCase("WriteOnlyObject", "Public Property Set WriteOnlyObject(ByVal arg As Collection)")]
        [TestCase("WriteOnlyVariant", "Public Property Let\\Set WriteOnlyVariant(ByVal arg As Variant)")]
        [TestCase("ReadOnlyValue", "Public Property Get ReadOnlyValue() As Long")]
        [TestCase("ReadOnlyObject", "Public Property Get ReadOnlyObject() As Collection")]
        [TestCase("ReadOnlyVariant", "Public Property Get ReadOnlyVariant() As Variant")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveableMemberSetPropertyDisplayNames(string member, string expectedDisplay)
        {
            var source =
$@"
Option Explicit

Public Property Let Fizz(arg As Long)
End Property

Public Property Get Fizz() As Long
End Property

Public Property Set Bizz(arg As Collection)
End Property

Public Property Get Bizz() As Collection
End Property

Public Property Set Dizz(arg As Variant)
End Property

Public Property Let Dizz(arg As Variant)
End Property

Public Property Get Dizz() As Variant
End Property

Public Property Set WriteOnlyObject(arg As Collection)
End Property

Public Property Let WriteOnlyValue(arg As Long)
End Property

Public Property Set WriteOnlyVariant(arg As Variant)
End Property

Public Property Let WriteOnlyVariant(arg As Variant)
End Property

Public Property Get ReadOnlyValue() As Long
End Property

Public Property Get ReadOnlyObject() As Collection
End Property

Public Property Get ReadOnlyVariant() As Variant
End Property
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, (member, DeclarationType.PropertyGet), sourceContent: source);
            var vbeStub = MoveMemberTestSupport.BuildVBEStub(moveDefinition, source);

            var displaySignature = MoveMemberTestSupport.ParseAndTest(vbeStub, ThisTest);

            StringAssert.AreEqualIgnoringCase(expectedDisplay, displaySignature);

            string ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var target = state.DeclarationFinder.MatchName(member);

                var model = new MoveMemberModel(target.First(), state as IDeclarationFinderProvider);
                var viewModel = new MoveableMemberSetViewModel(new MoveMemberViewModel(model), model.MoveableMemberSetByName(target.First().IdentifierName));
                return viewModel.ToDisplayString;
            }
        }

        [TestCase("Fizz", "Public Sub Fizz(arg1 As Long, arg2 As Collection)")]
        [TestCase("Bizz", "Public Function Bizz() As Long")]
        [TestCase("DIZZ", "Public Const DIZZ As Long = 45")]
        [TestCase("SAC", @"Private Const SAC As String = ""Sac""")]
        [TestCase("GIZZ", "Public GIZZ As Single")]
        [TestCase("HIZZ", "Public HIZZ As Variant")]
        [TestCase("LIZZ", "Public Const LIZZ As Long = OtherVal")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveableMemberSetNonPropertyDisplayNameX(string member, string expectedDisplay)
        {
            var source =
$@"
Option Explicit

Public Const DIZZ As Long = 45

Private Const SAD As String = ""Sad"", SAC As String = ""Sac"", OtherVal As Long = 75

Public Const LIZZ As Long = OtherVal

Public GIZZ As Single

Public HIZZ

Public Sub Fizz(arg1 As Long, arg2 As Collection)
End Sub

Public Function Bizz() As Long
End Function
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, (member, DeclarationType.PropertyGet), sourceContent: source);

            var vbeStub = MoveMemberTestSupport.BuildVBEStub(moveDefinition, source);
            var displaySignature = MoveMemberTestSupport.ParseAndTest(vbeStub, ThisTest);

            StringAssert.AreEqualIgnoringCase(expectedDisplay, displaySignature);


            string ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var target = state.DeclarationFinder.MatchName(member);

                var model = new MoveMemberModel(target.First(), state as IDeclarationFinderProvider);
                var viewModel = new MoveableMemberSetViewModel(new MoveMemberViewModel(model), model.MoveableMemberSetByName(target.First().IdentifierName));
                return viewModel.ToDisplayString;
            }
        }
    }
}
