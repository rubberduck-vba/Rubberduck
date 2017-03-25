using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Symbols
{
    [TestClass]
    public class SubroutineDeclarationTests
    {
        [TestMethod]
        public void SubroutinesHaveDeclarationTypeProcedure()
        {
            var subroutine = GetTestSub("testSub", null);

            Assert.IsTrue(subroutine.DeclarationType.HasFlag(DeclarationType.Procedure));
        }

            private static SubroutineDeclaration GetTestSub(string name, Attributes attributes)
            {
                var qualifiedName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new SubroutineDeclaration(qualifiedName, null, null, "test", Accessibility.Implicit, null, Selection.Home, true, null, attributes);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }


        [TestMethod]
        public void ByDefaultSubroutinesDoNotHaveParameters()
        {
            var subroutine = GetTestSub("testSub", null);

            Assert.IsFalse(subroutine.Parameters.Any());
        }


        [TestMethod]
        public void ParametersReturnsTheParametersAddedViaAddParameters()
        {
            var subroutine = GetTestSub("testSub", null);
            var inputParameter = GetTestParameter("testParameter", false, false, false);
            subroutine.AddParameter(inputParameter);
            var returnedParameter = subroutine.Parameters.SingleOrDefault();

            Assert.AreEqual(returnedParameter, inputParameter);
        }

            private static ParameterDeclaration GetTestParameter(string name, bool isOptional, bool isByRef, bool isParamArray)
            {
                var qualifiedParameterName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ParameterDeclaration(qualifiedParameterName, null, "test", null, "test", isOptional, isByRef, false, isParamArray);
            }


        [TestMethod]
        public void ByDefaultSubroutinesAreNotDefaultMembers()
        {
            var subroutine = GetTestSub("testSub", null);

            Assert.IsFalse(subroutine.IsDefaultMember);
        }


        [TestMethod]
        public void SubroutinesAreDefaultMembersIfTheyHaveTheDefaultMemberAttribute()
        {
            var attributes = new Attributes();
            attributes.AddDefaultMemberAttribute("testSub");
            var subroutine = GetTestSub("testSub", attributes);

            Assert.IsTrue(subroutine.IsDefaultMember);
        }

    }
}
