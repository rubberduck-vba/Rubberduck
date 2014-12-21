using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;

namespace RubberduckTests
{
    [TestClass]
    public class CodeInspectionTests
    {
        private readonly IEnumerable<ISyntax> _grammar = Assembly.GetAssembly(typeof(ISyntax))
                                  .GetTypes()
                                  .Where(type => type.BaseType == typeof(SyntaxBase))
                                  .Select(type =>
                                  {
                                      var constructorInfo = type.GetConstructor(Type.EmptyTypes);
                                      return constructorInfo != null ? constructorInfo.Invoke(Type.EmptyTypes) : null;
                                  })
                                  .Cast<ISyntax>()
                                  .ToList();
        
        [TestMethod]
        public void OptionExplicitInspectionTest()
        {
            var code = "Sub Foo()\n\nEnd Sub";
            var parser = new Parser(_grammar);
            var node = parser.Parse("CodeInspectionTests", "TestModule", code, false);
            var projectNode = new ProjectNode("CodeInspectionTests", new[] { node });

            var inspection = new OptionExplicitInspection();
            var results = inspection.GetInspectionResults(projectNode);

            Assert.IsTrue(results.Any(result => result.GetType() == typeof(OptionExplicitInspectionResult)));
        }

        [TestMethod]
        public void VariableTypeNotDeclaredInspectionTest()
        {
            var code = "Private foo\n";
            var parser = new Parser(_grammar);
            var node = parser.Parse("CodeInspectionTests", "TestModule", code, false);
            var projectNode = new ProjectNode("CodeInspectionTests", new[] { node });

            var inspection = new VariableTypeNotDeclaredInspection();
            var results = inspection.GetInspectionResults(projectNode);

            Assert.IsTrue(results.Any(result => result.GetType() == typeof(VariableTypeNotDeclaredInspectionResult)));
        }

        [TestMethod]
        public void ObsoleteCommentSyntaxInspectionTest()
        {
            var code = "Rem obsolete syntax for a comment\n";
            var parser = new Parser(_grammar);
            var node = parser.Parse("CodeInspectionTests", "TestModule", code, false);
            var projectNode = new ProjectNode("CodeInspectionTests", new[] { node });

            var inspection = new ObsoleteCommentSyntaxInspection();
            var results = inspection.GetInspectionResults(projectNode);

            Assert.IsTrue(results.Any(result => result.GetType() == typeof(ObsoleteCommentSyntaxInspectionResult)));
        }

        [TestMethod]
        public void ImplicitVariantReturnTypeInspectionTest()
        {
            var code = "Private Function Foo(ByVal bar As Integer)\n\nEnd Function";
            var parser = new Parser(_grammar);
            var node = parser.Parse("CodeInspectionTests", "TestModule", code, false);
            var projectNode = new ProjectNode("CodeInspectionTests", new[] { node });

            var inspection = new ImplicitVariantReturnTypeInspection();
            var results = inspection.GetInspectionResults(projectNode);

            Assert.IsTrue(results.Any(result => result.GetType() == typeof(ImplicitVariantReturnTypeInspectionResult)));
        }

        [TestMethod]
        public void ImplicitVariantReturnTypeInspectionTestWithParamArray()
        {
            var code = "Private Function Foo(ByVal bar As Integer, ParamArray buzz())\n\nEnd Function";
            var parser = new Parser(_grammar);
            var node = parser.Parse("CodeInspectionTests", "TestModule", code, false);
            var projectNode = new ProjectNode("CodeInspectionTests", new[] { node });

            var inspection = new ImplicitVariantReturnTypeInspection();
            var results = inspection.GetInspectionResults(projectNode);

            Assert.IsTrue(results.Any(result => result.GetType() == typeof(ImplicitVariantReturnTypeInspectionResult)));
        }

        [TestMethod]
        public void ImplicitVariantReturnTypeInspectionStressTestWithParamArray()
        {
            var code = "Private Function Foo(ByVal bar As Integer, ParamArray buzz()) As Variant()\n\nEnd Function";
            var parser = new Parser(_grammar);
            var node = parser.Parse("CodeInspectionTests", "TestModule", code, false);
            var projectNode = new ProjectNode("CodeInspectionTests", new[] { node });

            var inspection = new ImplicitVariantReturnTypeInspection();
            var results = inspection.GetInspectionResults(projectNode);

            Assert.IsFalse(results.Any(result => result.GetType() == typeof(ImplicitVariantReturnTypeInspectionResult)));
        }

        [TestMethod]
        public void ImplicitByRefParameterInspectionTest()
        {
            var code = "Private Sub Foo(bar As Integer)\n\nEnd Sub";
            var parser = new Parser(_grammar);
            var node = parser.Parse("CodeInspectionTests", "TestModule", code, false);
            var projectNode = new ProjectNode("CodeInspectionTests", new[] { node });

            var inspection = new ImplicitByRefParameterInspection();
            var results = inspection.GetInspectionResults(projectNode);

            Assert.IsTrue(results.Any(result => result.GetType() == typeof(ImplicitByRefParameterInspectionResult)));
        }
    }
}
