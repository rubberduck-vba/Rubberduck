using NUnit.Framework;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class PropertyGetDeclarationTests
    {
        [Test]
        [Category("Resolver")]
        public void PropertyGetsHaveDeclarationTypePropertyGet()
        {
            var propertyGet = GetTestPropertyGet("test", null);

            Assert.IsTrue(propertyGet.DeclarationType.HasFlag(DeclarationType.PropertyGet));
        }

            private static PropertyGetDeclaration GetTestPropertyGet(string name, Attributes attributes)
            {
                var qualifiedName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new PropertyGetDeclaration(qualifiedName, null, null, "test", null, "test", Accessibility.Implicit, null, null, Selection.Home, false, true, null, attributes);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }


        [Test]
        [Category("Resolver")]
        public void ByDefaultPropertyGetsDoNotHaveParameters()
        {
            var propertyGet = GetTestPropertyGet("test", null);

            Assert.IsFalse(propertyGet.Parameters.Any());
        }


        [Test]
        [Category("Resolver")]
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


        [Test]
        [Category("Resolver")]
        public void ByDefaultPropertyGetsAreNotDefaultMembers()
        {
            var propertyGet = GetTestPropertyGet("test", null);

            Assert.IsFalse(propertyGet.IsDefaultMember);
        }


        [Test]
        [Category("Resolver")]
        public void PropertyGetsAreDefaultMembersIfTheyHaveTheDefaultMemberAttribute()
        {
            var attributes = new Attributes();
            attributes.AddDefaultMemberAttribute("test");
            var propertyGet = GetTestPropertyGet("test", attributes);

            Assert.IsTrue(propertyGet.IsDefaultMember);
        }
    }
}
