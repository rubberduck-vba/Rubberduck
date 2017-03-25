using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace RubberduckTests.Symbols
{
    [TestClass]
    public class ExternalProcedureDeclarationTests
    {
        [TestMethod]
        public void ByDefaultExternalProceduresDoNotHaveParameters()
        {
            var externalProcedure = GetTestExternalProcedure("testProcedure");

            Assert.IsFalse(externalProcedure.Parameters.Any());
        }

            private static ExternalProcedureDeclaration GetTestExternalProcedure(string name)
            {
                var qualifiedProcedureName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ExternalProcedureDeclaration(qualifiedProcedureName, null, null, DeclarationType.Procedure, "test", null, Accessibility.Public, null, Selection.Home, true, null);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }


        [TestMethod]
        public void ParametersReturnsTheParametersAddedViaAddParameters()
        {
            var externalProcedure = GetTestExternalProcedure("testProcedure");
            var inputParameter = GetTestParameter("testParameter", false, false, false);
            externalProcedure.AddParameter(inputParameter);
            var returnedParameter = externalProcedure.Parameters.SingleOrDefault();

            Assert.AreEqual(returnedParameter, inputParameter);
        }

            private static ParameterDeclaration GetTestParameter(string name, bool isOptional, bool isByRef, bool isParamArray)
            {
                var qualifiedParameterName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ParameterDeclaration(qualifiedParameterName, null, "test", null, "test", isOptional, isByRef, false, isParamArray);
            }

    }
}
