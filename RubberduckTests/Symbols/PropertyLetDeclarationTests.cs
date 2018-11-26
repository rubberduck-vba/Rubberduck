using NUnit.Framework;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class PropertyLetDeclarationTests
    {
        [Test]
        [Category("Resolver")]
        public void PropertyLetsHaveDeclarationTypePropertyLet()
        {
            var propertyLet = GetTestPropertyLet("test", null);

            Assert.IsTrue(propertyLet.DeclarationType.HasFlag(DeclarationType.PropertyLet));
        }

            private static PropertyLetDeclaration GetTestPropertyLet(string name, Attributes attributes)
            {
                var qualifiedName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new PropertyLetDeclaration(qualifiedName, null, null, "test", Accessibility.Implicit, null, null, Selection.Home, true, null, attributes);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }


        [Test]
        [Category("Resolver")]
        public void ByDefaultPropertyLetsDoNotHaveParameters()
        {
            var propertyLet = GetTestPropertyLet("test", null);

            Assert.IsFalse(propertyLet.Parameters.Any());
        }


        [Test]
        [Category("Resolver")]
        public void ParametersReturnsTheParametersAddedViaAddParameters()
        {
            var propertyLet = GetTestPropertyLet("test", null);
            var inputParameter = GetTestParameter("testParameter", false, false, false);
            propertyLet.AddParameter(inputParameter);
            var returnedParameter = propertyLet.Parameters.SingleOrDefault();

            Assert.AreEqual(returnedParameter, inputParameter);
        }

            private static ParameterDeclaration GetTestParameter(string name, bool isOptional, bool isByRef, bool isParamArray)
            {
                var qualifiedParameterName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ParameterDeclaration(qualifiedParameterName, null, "test", null, "test", isOptional, isByRef, false, isParamArray);
            }


        [Test]
        [Category("Resolver")]
        public void ByDefaultPropertyLetsAreNotDefaultMembers()
        {
            var propertyLet = GetTestPropertyLet("test", null);

            Assert.IsFalse(propertyLet.IsDefaultMember);
        }


        [Test]
        [Category("Resolver")]
        public void PropertyLetsAreDefaultMembersIfTheyHaveTheDefaultMemberAttribute()
        {
            var attributes = new Attributes();
            attributes.AddDefaultMemberAttribute("test");
            var propertyLet = GetTestPropertyLet("test", attributes);

            Assert.IsTrue(propertyLet.IsDefaultMember);
        }
    }
}
