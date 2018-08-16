using System.Linq;
using NUnit.Framework;
using Rubberduck.UnitTesting;
using RubberduckTests.Mocks;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    public class TestMethodTests
    {
        [Test]
        public void TestCategoryIsAssigned()
        {
            const string code = @"
'@TestMethod(""Category"")
Sub Foo()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var testMethodDeclaration = state.AllUserDeclarations.First(declaration => declaration.IdentifierName == "Foo");
                var testMethod = new TestMethod(testMethodDeclaration);

                Assert.AreEqual("Category", testMethod.Category.Name);
            }
        }

        [Test]
        public void TestCategoryIsEmptyWhenNoCategorySpecified()
        {
            const string code = @"
'@TestMethod
Sub Foo()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var testMethodDeclaration = state.AllUserDeclarations.First(declaration => declaration.IdentifierName == "Foo");
                var testMethod = new TestMethod(testMethodDeclaration);

                Assert.AreEqual("", testMethod.Category.Name);
            }
        }

        [Test]
        public void TestCategoryIsEmptyWhenSpecifiedCategoryHasWhiteSpaceOnly()
        {
            const string code = @"
'@TestMethod(""   "")
Sub Foo()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var testMethodDeclaration = state.AllUserDeclarations.First(declaration => declaration.IdentifierName == "Foo");
                var testMethod = new TestMethod(testMethodDeclaration);

                Assert.AreEqual("", testMethod.Category.Name);
            }
        }
    }
}
