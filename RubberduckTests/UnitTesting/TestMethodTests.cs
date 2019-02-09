using System.Linq;
using NUnit.Framework;
using Rubberduck.Resources.UnitTesting;
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
        public void TestCategoryIsUncategorizedWhenNoCategorySpecified()
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

                Assert.AreEqual(TestExplorer.TestExplorer_Uncategorized, testMethod.Category.Name);
            }
        }

        [Test]
        public void TestCategoryIsUncategorizedWhenSpecifiedCategoryHasWhiteSpaceOnly()
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

                Assert.AreEqual(TestExplorer.TestExplorer_Uncategorized, testMethod.Category.Name);
            }
        }
    }
}
