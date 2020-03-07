using NUnit.Framework;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    [Category("Inspections")]
    [Category("UnderscoreInPublicMember")]
    public class UnderscoreInPublicClassModuleMemberInspectionTests : InspectionTestsBase
    {
        [Test]
        public void BasicExample_Sub()
        {
            const string inputCode =
                @"Public Sub Test_This_Out()
End Sub";
            Assert.AreEqual(1, InspectionResultsForModules(("TestClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        public void Basic_Ignored()
        {
            const string inputCode =
                @"'@Ignore UnderscoreInPublicClassModuleMember
Public Sub This_Is_Ignored()
End Sub

Public Sub This_Should_Be_Marked()
End Sub";
            Assert.AreEqual(1, InspectionResultsForModules(("TestClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        public void Basic_IgnoreModule()
        {
            const string inputCode =
                @"'@IgnoreModule UnderscoreInPublicClassModuleMember
Public Sub This_Is_Ignored()
End Sub

Public Sub This_Is_Also_Ignored()
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        public void BasicExample_Function()
        {
            const string inputCode =
                @"Public Function Test_This_Out() As Integer
End Function";
            Assert.AreEqual(1, InspectionResultsForModules(("TestClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        public void BasicExample_Property()
        {
            const string inputCode =
                @"Public Property Get Test_This_Out() As Integer
End Property";
            Assert.AreEqual(1, InspectionResultsForModules(("TestClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        public void StandardModule()
        {
            const string inputCode =
                @"Public Sub Test_This_Out()
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        public void NoUnderscore()
        {
            const string inputCode =
                @"Public Sub Foo()
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        public void FriendMember_WithUnderscore()
        {
            const string inputCode =
                @"Friend Sub Test_This_Out()
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        public void Implicit_WithUnderscore()
        {
            const string inputCode =
                @"Sub Test_This_Out()
End Sub";
            Assert.AreEqual(1, InspectionResultsForModules(("TestClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        public void ImplementsInterface()
        {
            const string inputCode1 =
       @"Public Sub Foo()
End Sub";

            const string inputCode2 =
                @"Implements Class1

Public Sub Class1_Foo()
    Err.Raise 5 'TODO implement interface member
End Sub
";
            var modules = new(string, string, ComponentType)[] 
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
            };
            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new UnderscoreInPublicClassModuleMemberInspection(state);
        }
    }
}
