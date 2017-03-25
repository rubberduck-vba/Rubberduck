using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Symbols
{
    [TestClass]
    public class PropertySetDeclarationTests
    {
        [TestMethod]
        public void PropertySetsHaveDeclarationTypePropertySet()
        {
            var propertySet = GetTestPropertySet("test", null);

            Assert.IsTrue(propertySet.DeclarationType.HasFlag(DeclarationType.PropertySet));
        }

            private static PropertySetDeclaration GetTestPropertySet(string name, Attributes attributes)
            {
                var qualifiedName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new PropertySetDeclaration(qualifiedName, null, null, "test", Accessibility.Implicit, null, Selection.Home, true, null, attributes);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }


        [TestMethod]
        public void ByDefaultPropertySetsDoNotHaveParameters()
        {
            var propertySet = GetTestPropertySet("test", null);

            Assert.IsFalse(propertySet.Parameters.Any());
        }


        [TestMethod]
        public void ParametersReturnsTheParametersAddedViaAddParameters()
        {
            var propertySet = GetTestPropertySet("test", null);
            var inputParameter = GetTestParameter("testParameter", false, false, false);
            propertySet.AddParameter(inputParameter);
            var returnedParameter = propertySet.Parameters.SingleOrDefault();

            Assert.AreEqual(returnedParameter, inputParameter);
        }

            private static ParameterDeclaration GetTestParameter(string name, bool isOptional, bool isByRef, bool isParamArray)
            {
                var qualifiedParameterName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ParameterDeclaration(qualifiedParameterName, null, "test", null, "test", isOptional, isByRef, false, isParamArray);
            }


        [TestMethod]
        public void ByDefaultPropertySetsAreNotDefaultMembers()
        {
            var propertySet = GetTestPropertySet("test", null);

            Assert.IsFalse(propertySet.IsDefaultMember);
        }


        [TestMethod]
        public void PropertySetsAreDefaultMembersIfTheyHaveTheDefaultMemberAttribute()
        {
            var attributes = new Attributes();
            attributes.AddDefaultMemberAttribute("test");
            var propertySet = GetTestPropertySet("test", attributes);

            Assert.IsTrue(propertySet.IsDefaultMember);
        }
    }
}
