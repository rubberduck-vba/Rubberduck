using System.Linq;
using Antlr4.Runtime;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.TypeResolvers;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.ExpressionResolving
{
    [TestFixture]
    public class SetTypeResolverTests
    {
        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("TestProject.Class1", "TestProject.Class1")]
        [TestCase("Variant", "Variant")]
        [TestCase("Object", "Object")]
        [TestCase("Long", null)]
        public void SimpleNameExpression_SetTypeNameTests(string typeName, string expectedSetTypeName)
        {
            const string class1 =
                @"
Private Sub Foo()
End Sub
";

            var module1 =
                $@"
Private Sub Bar()
    Dim cls As {typeName}
    Set cls = cls
End Sub
";

            var expressionSelection = new Selection(4, 15, 4, 18);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("TestProject.Class1", "TestProject.Class1")]
        [TestCase("Variant", null)]
        [TestCase("Object", null)]
        [TestCase("Long", null)]
        public void SimpleNameExpression_SetTypeDeclarationTests(string typeName, string expectedNameOfSetTypeDeclaration)
        {
            const string class1 =
                @"
Private Sub Foo()
End Sub
";

            var module1 =
                $@"
Private Sub Bar()
    Dim cls As {typeName}
    Set cls = cls
End Sub
";

            var expressionSelection = new Selection(4, 15, 4, 18);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        public void InstanceExpression_SetTypeName_ReturnsNameOfContainingClass()
        {
            const string class1 =
                @"
Private Sub Foo()
    Dim bar As Variant
    Set bar = Me
End Sub
";

            var expressionSelection = new Selection(4, 15, 4, 17);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Class1", expressionSelection);
            var expectedSetTypeName = "TestProject.Class1";

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        public void InstanceExpression_SetTypeDeclaration_ReturnsDeclarationOfContainingClass()
        {
            const string class1 =
                @"
Private Sub Foo()
    Dim bar As Variant
    Set bar = Me
End Sub
";

            var expressionSelection = new Selection(4, 15, 4, 17);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Class1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();
            var expectedNameOfSetTypeDeclaration = "TestProject.Class1";

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }


        private Declaration ExpressionTypeDeclaration(IVBE vbe, string componentName, Selection selection)
        {
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var resolver = ExpressionResolverUnderTest(state);
                var module = state.DeclarationFinder.AllModules.First(qmn => qmn.ComponentName.Equals(componentName));
                var expression = TestExpression(state, module, selection);
                return resolver.SetTypeDeclaration(expression, module);
            }
        }

        private string ExpressionTypeName(IVBE vbe, string componentName, Selection selection)
        {
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var resolver = ExpressionResolverUnderTest(state);
                var module = state.DeclarationFinder.AllModules.First(qmn => qmn.ComponentName.Equals(componentName));
                var expression = TestExpression(state, module, selection);
                return resolver.SetTypeName(expression, module);
            }
        }

        private VBAParser.ExpressionContext TestExpression(IParseTreeProvider parseTreeProvider, QualifiedModuleName module, Selection selection)
        {
            if (!(parseTreeProvider.GetParseTree(module, CodeKind.CodePaneCode) is ParserRuleContext context))
            {
                return null;
            }

            if (!context.GetSelection().Contains(selection))
            {
                return null;
            }

            return TestExpression(context, selection);
        }

        private VBAParser.ExpressionContext TestExpression(ParserRuleContext context, Selection selection)
        {
            if (context == null)
            {
                return null;
            }

            if (context is VBAParser.ExpressionContext expression)
            {
                return expression;
            }

            if (context.children == null)
            {
                return null;
            }

            foreach (var child in context.children)
            {
                if (child is ParserRuleContext childContext && childContext.GetSelection().Contains(selection))
                {
                    return TestExpression(childContext, selection);
                }
            }

            return null;
        }

        private static ISetTypeResolver ExpressionResolverUnderTest(IDeclarationFinderProvider declarationFinderProvider)
        {
            return new SetTypeResolver(declarationFinderProvider);
        }
    }
}