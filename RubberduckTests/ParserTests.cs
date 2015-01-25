using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;

namespace RubberduckTests
{
    [TestClass]
    public class ParserTests
    {
        [TestMethod]
        public void UnspecifiedProcedureVisibilityIsImplicit()
        {
            IRubberduckParser parser = new VBParser();
            var code = "Sub Foo()\n    Dim bar As Integer\nEnd Sub";

            var module = parser.Parse("project", "component", code);
            var procedure = (ProcedureNode)module.Children.First();

            Assert.AreEqual(procedure.Accessibility, VBAccessibility.Implicit);
        }

        [TestMethod]
        public void UnspecifiedReturnTypeGetsFlagged()
        {
            IRubberduckParser parser = new VBParser();
            var code = "Function Foo()\n    Dim bar As Integer\nEnd Function";

            var module = parser.Parse("project", "component", code);
            var procedure = (ProcedureNode)module.Children.First();

            Assert.AreEqual(procedure.ReturnType, "Variant");            
            Assert.IsTrue(procedure.IsImplicitReturnType);
        }

        [TestMethod]
        public void LocalDimMakesPrivateVariable()
        {
            IRubberduckParser parser = new VBParser();
            var code = "Sub Foo()\n    Dim bar As Integer\nEnd Sub";

            var module = parser.Parse("project", "component", code);
            var procedure = module.Children.First();
            var declaration = procedure.Children.First();
            var variable = (VariableNode)declaration.Children.First();

            Assert.AreEqual(variable.Accessibility, VBAccessibility.Private);
        }

        [TestMethod]
        public void TypeHintsGetFlagged()
        {
            IRubberduckParser parser = new VBParser();
            var code = "Sub Foo()\n    Dim bar$\nEnd Sub";

            var module = parser.Parse("project", "component", code);
            var procedure = module.Children.First();
            var declaration = procedure.Children.First();
            var variable = (VariableNode)declaration.Children.First();

            Assert.IsTrue(variable.IsUsingTypeHint);
        }

        [TestMethod]
        public void ImplicitTypeGetsFlagged()
        {
            IRubberduckParser parser = new VBParser();
            var code = "Sub Foo()\n    Dim bar\nEnd Sub";

            var module = parser.Parse("project", "component", code);
            var procedure = module.Children.First();
            var declaration = procedure.Children.First();
            var variable = (VariableNode)declaration.Children.First();

            Assert.IsTrue(variable.IsImplicitlyTyped);
        }
    }
}
