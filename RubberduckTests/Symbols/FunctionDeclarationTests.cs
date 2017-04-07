using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Symbols
{
    [TestClass]
    public class FunctionDeclarationTests
    {
        [TestMethod]
        public void FunctionsHaveDeclarationTypeFunction()
        {
            var function = GetTestFunction("testFoo", null);

            Assert.IsTrue(function.DeclarationType.HasFlag(DeclarationType.Function));
        }

            private static FunctionDeclaration GetTestFunction(string name, Attributes attributes)
            {
                var qualifiedName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new FunctionDeclaration(qualifiedName, null, null, "test", null, "test", Accessibility.Implicit, null, Selection.Home, false, true, null, attributes);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }


        [TestMethod]
        public void ByDefaultFunctionsDoNotHaveParameters()
        {
            var function = GetTestFunction("testFoo", null);

            Assert.IsFalse(function.Parameters.Any());
        }


        [TestMethod]
        public void ParametersReturnsTheParametersAddedViaAddParameters()
        {
            var function = GetTestFunction("testFoo", null);
            var inputParameter = GetTestParameter("testParameter", false, false, false);
            function.AddParameter(inputParameter);
            var returnedParameter = function.Parameters.SingleOrDefault();

            Assert.AreEqual(returnedParameter, inputParameter);
        }

            private static ParameterDeclaration GetTestParameter(string name, bool isOptional, bool isByRef, bool isParamArray)
            {
                var qualifiedParameterName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ParameterDeclaration(qualifiedParameterName, null, "test", null, "test", isOptional, isByRef, false, isParamArray);
            }


        [TestMethod]
        public void ByDefaultFunctionsAreNotDefaultMembers()
        {
            var function = GetTestFunction("testFoo", null);

            Assert.IsFalse(function.IsDefaultMember);
        }


        [TestMethod]
        public void FunctionsAreDefaultMembersIfTheyHaveTheDefaultMemberAttribute()
        {
            var attributes = new Attributes();
            attributes.AddDefaultMemberAttribute("testFoo");
            var function = GetTestFunction("testFoo", attributes);

            Assert.IsTrue(function.IsDefaultMember);
        }
    }
}
