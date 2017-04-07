using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Symbols
{
    [TestClass]
    public class PropertyGetDeclarationTests
    {
        [TestMethod]
        public void PropertyGetsHaveDeclarationTypePropertyGet()
        {
            var propertyGet = GetTestPropertyGet("test", null);

            Assert.IsTrue(propertyGet.DeclarationType.HasFlag(DeclarationType.PropertyGet));
        }

            private static PropertyGetDeclaration GetTestPropertyGet(string name, Attributes attributes)
            {
                var qualifiedName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new PropertyGetDeclaration(qualifiedName, null, null, "test", null, "test", Accessibility.Implicit, null, Selection.Home, false, true, null, attributes);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }


        [TestMethod]
        public void ByDefaultPropertyGetsDoNotHaveParameters()
        {
            var propertyGet = GetTestPropertyGet("test", null);

            Assert.IsFalse(propertyGet.Parameters.Any());
        }


        [TestMethod]
        public void ParametersReturnsTheParametersAddedViaAddParameters()
        {
            var propertyGet = GetTestPropertyGet("test", null);
            var inputParameter = GetTestParameter("testParameter", false, false, false);
            propertyGet.AddParameter(inputParameter);
            var returnedParameter = propertyGet.Parameters.SingleOrDefault();

            Assert.AreEqual(returnedParameter, inputParameter);
        }

            private static ParameterDeclaration GetTestParameter(string name, bool isOptional, bool isByRef, bool isParamArray)
            {
                var qualifiedParameterName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ParameterDeclaration(qualifiedParameterName, null, "test", null, "test", isOptional, isByRef, false, isParamArray);
            }


        [TestMethod]
        public void ByDefaultPropertyGetsAreNotDefaultMembers()
        {
            var propertyGet = GetTestPropertyGet("test", null);

            Assert.IsFalse(propertyGet.IsDefaultMember);
        }


        [TestMethod]
        public void PropertyGetsAreDefaultMembersIfTheyHaveTheDefaultMemberAttribute()
        {
            var attributes = new Attributes();
            attributes.AddDefaultMemberAttribute("test");
            var propertyGet = GetTestPropertyGet("test", attributes);

            Assert.IsTrue(propertyGet.IsDefaultMember);
        }
    }
}
