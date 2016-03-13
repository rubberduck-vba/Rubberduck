using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace RubberduckTests.Grammar
{
    [TestClass]
    public class ResolverTests
    {
        [TestMethod]
        public void FunctionReturnValueAssignment_IsReferenceToFunctionDeclaration()
        {
            var code = @"
Public Function Foo() As String
    Foo = 42
End Function
";

        }
    }
}