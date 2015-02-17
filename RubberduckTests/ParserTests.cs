using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Antlr4.Runtime.Tree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

namespace RubberduckTests
{
    [TestClass]
    public class ParserTests
    {
        [TestMethod]
        public void GetPublicProceduresReturnsPublicSubs()
        {
            IRubberduckParser parser = new RubberduckParser();
            var code = "Sub Foo()\nEnd Sub\n\nPrivate Sub FooBar()\nEnd Sub\n\nPublic Sub Bar()\nEnd Sub\n\nPublic Sub BarFoo(ByVal fb As Long)\nEnd Sub\n\nFunction GetFoo() As Bar\nEnd Function";

            var module = parser.Parse(code);
            var walker = new ParseTreeWalker();

            var listener = new PublicSubListener();
            walker.Walk(listener, module);
            var procedures = listener.Members.ToList();

            var parameterless = procedures
                .Where(p => p.argList().Arg().Count == 0);

            Assert.AreEqual(3, procedures.Count);
            Assert.AreEqual(2, parameterless.Count());
        }

        [TestMethod]
        public void UnspecifiedProcedureVisibilityIsImplicit()
        {
            IRubberduckParser parser = new RubberduckParser();
            var code = "Sub Foo()\r\n    Dim bar As Integer\r\nEnd Sub";

            var module = parser.Parse("project", "component", code);
            var procedure = (ProcedureNode)module.Children.First();

            Assert.AreEqual(procedure.Accessibility, VBAccessibility.Implicit);
        }

        [TestMethod]
        public void DeclarationSectionConst()
        {
            IRubberduckParser parser = new RubberduckParser();
            var code = "Const foo As Integer = 42\n\nOption Explicit\n\n";

            var module = parser.Parse("project", "component", code);
            var declaration = (ConstDeclarationNode)module.Children.First();
            var constant = (DeclaredIdentifierNode)declaration.Children.First();

            Assert.AreEqual(constant.Name, "foo");
        }

        [TestMethod]
        public void UnspecifiedReturnTypeGetsFlagged()
        {
            IRubberduckParser parser = new RubberduckParser();
            var code = "Function Foo()\n    Dim bar As Integer\nEnd Function";

            var module = parser.Parse("project", "component", code);
            var procedure = (ProcedureNode)module.Children.First();
            
            Assert.AreEqual(procedure.ReturnType, "Variant");            
            Assert.IsTrue(procedure.IsImplicitReturnType);
        }

        [TestMethod]
        public void LocalDimMakesPrivateVariable()
        {
            IRubberduckParser parser = new RubberduckParser();
            var code = "Sub Foo()\n    Dim bar As Integer\nEnd Sub";

            var module = parser.Parse("project", "component", code);
            var procedure = module.Children.First();
            var declaration = procedure.Children.First();
            var variable = (DeclaredIdentifierNode)declaration.Children.First();

            Assert.AreEqual(variable.Accessibility, VBAccessibility.Private);
        }

        [TestMethod]
        public void TypeHintsGetFlagged()
        {
            IRubberduckParser parser = new RubberduckParser();
            var code = "Sub Foo()\n    Dim bar$\nEnd Sub";

            var module = parser.Parse("project", "component", code);
            var procedure = module.Children.First();
            var declaration = procedure.Children.First();
            var variable = (DeclaredIdentifierNode)declaration.Children.First();

            Assert.AreEqual(variable.TypeName, "String");
            Assert.IsTrue(variable.IsUsingTypeHint);
        }

        [TestMethod]
        public void ImplicitTypeGetsFlagged()
        {
            IRubberduckParser parser = new RubberduckParser();
            var code = "Sub Foo()\n    Dim bar\nEnd Sub";

            var module = parser.Parse("project", "component", code);
            var procedure = module.Children.First();
            var declaration = procedure.Children.First();
            var variable = (DeclaredIdentifierNode)declaration.Children.First();

            Assert.IsTrue(variable.IsImplicitlyTyped);
        }
    }
}
