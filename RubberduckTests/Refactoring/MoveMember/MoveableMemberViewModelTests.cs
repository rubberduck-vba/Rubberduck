using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.UI.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

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
            var result = GenerateDisplayString(member, (MoveEndpoints.StdToStd.SourceModuleName(), source, ComponentType.StandardModule)); // moveDefinition/*, source*/);

            StringAssert.AreEqualIgnoringCase(expectedDisplay, result);
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
        public void MoveableMemberSetNonPropertyDisplayNames(string member, string expectedDisplay)
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
            var result = GenerateDisplayString(member, MoveEndpoints.StdToStd.ToSourceTuple(source));

            StringAssert.AreEqualIgnoringCase(expectedDisplay, result);
        }

        public static string GenerateDisplayString(string targetIdentifier, params (string, string, ComponentType)[] modules)
        {
            var vbeStub = MockVbeBuilder.BuildFromModules(modules).Object;
            (RubberduckParserState state, IRewritingManager rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbeStub);
            using (state)
            {
                var resolver = new MoveMemberTestsResolver(state, rewritingManager);

                var targets = state.DeclarationFinder.MatchName(targetIdentifier);
                var moveableMemberSet = resolver.Resolve<MoveableMemberSetFactory>()
                    .Create(targets);

                var viewModel = new MoveableMemberSetViewModel(vm => { }, moveableMemberSet);
                return viewModel.ToDisplayString;
            }
        }
    }
}
