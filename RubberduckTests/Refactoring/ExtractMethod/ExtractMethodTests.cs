using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractMethod;

namespace RubberduckTests.Refactoring.ExtractMethod
{

    [TestClass]
    public class ExtractedMethodTests
    {
        [TestClass]
        public class WhenAMethodIsDefined: ExtractedMethodTests
        {

            [TestCategory("ExtractedMethodTests")]
            [TestMethod]
            public void shouldReturnStringCorrectly()
            {
                var method = new ExtractedMethod();
                method.Accessibility = Accessibility.Private;
                method.MethodName = "Bar";
                method.ReturnValue = null;
                var insertCode = "Bar x";
                var newParam = new ExtractedParameter("Integer", ExtractedParameter.PassedBy.ByVal, "x");
                method.Parameters = new List<ExtractedParameter>() { newParam };

                var actual = method.NewMethodCall();
                Debug.Print(method.NewMethodCall());
                
                Assert.AreEqual(insertCode, actual);


            }
        }
    }

}
