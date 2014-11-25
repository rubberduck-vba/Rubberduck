using System;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;

namespace RubberduckTests
{
    [TestClass]
    public class ParserTests
    {
        [TestMethod]
        public void TestSimpleDeclaration()
        {
            var parser = new Parser();
            const string code = "Dim foo As Integer\n";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var declaration = result.ChildNodes.FirstOrDefault() as DeclarationNode;
            Assert.IsNotNull(declaration);

            var identifier = declaration.ChildNodes.FirstOrDefault() as IdentifierNode;
            Assert.IsNotNull(identifier);

            Assert.AreEqual("foo", identifier.Name);
            Assert.AreEqual("Integer", identifier.TypeName);
            Assert.IsTrue(identifier.IsTypeSpecified);
        }

        [TestMethod]
        public void TestDeclarationWithAssignment()
        {
            var parser = new Parser();
            const string code = "Private Strings As New StringType\n";

            var match = Regex.Match(code, VBAGrammar.GeneralDeclarationSyntax);
            Assert.IsTrue(match.Success);

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var declaration = result.ChildNodes.FirstOrDefault() as DeclarationNode;
            Assert.IsNotNull(declaration);

            var identifier = declaration.ChildNodes.FirstOrDefault() as IdentifierNode;
            Assert.IsNotNull(identifier);

            Assert.AreEqual("Strings", identifier.Name);
            Assert.AreEqual("StringType", identifier.TypeName);
            Assert.IsTrue(identifier.IsTypeSpecified);
            Assert.IsTrue(identifier.IsInitialized);
        }

        [TestMethod]
        public void TestSimpleArrayDeclaration()
        {
            var parser = new Parser();
            const string code = "Dim foo() As Integer\n";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var declaration = result.ChildNodes.FirstOrDefault() as DeclarationNode;
            Assert.IsNotNull(declaration);

            var identifier = declaration.ChildNodes.FirstOrDefault() as IdentifierNode;
            Assert.IsNotNull(identifier);

            Assert.AreEqual("foo", identifier.Name);
            Assert.AreEqual("Integer", identifier.TypeName);
            Assert.IsTrue(identifier.IsTypeSpecified);
            Assert.IsTrue(identifier.IsArray);
        }

        [TestMethod]
        public void TestInitializedArrayDeclaration()
        {
            var parser = new Parser();
            const string code = "Dim foo(1 To 10) As Integer\n";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var declaration = result.ChildNodes.FirstOrDefault() as DeclarationNode;
            Assert.IsNotNull(declaration);

            var identifier = declaration.ChildNodes.FirstOrDefault() as IdentifierNode;
            Assert.IsNotNull(identifier);

            Assert.AreEqual("foo", identifier.Name);
            Assert.AreEqual("Integer", identifier.TypeName);
            Assert.IsTrue(identifier.IsTypeSpecified);
            Assert.IsTrue(identifier.IsArray);
        }

        [TestMethod]
        public void TestMultiDimArrayDeclaration()
        {
            var parser = new Parser();
            const string code = "Dim foo(1 To 10, 1 To 5) As Integer\n";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var declaration = result.ChildNodes.FirstOrDefault() as DeclarationNode;
            Assert.IsNotNull(declaration);

            var identifier = declaration.ChildNodes.FirstOrDefault() as IdentifierNode;
            Assert.IsNotNull(identifier);

            Assert.AreEqual("foo", identifier.Name);
            Assert.AreEqual("Integer", identifier.TypeName);
            Assert.IsTrue(identifier.IsTypeSpecified);
            Assert.IsTrue(identifier.IsArray);
            Assert.AreEqual(2, identifier.ArrayDimensionsCount);
        }

        [TestMethod]
        public void TestTypeSpecifierSimpleDeclaration()
        {
            var parser = new Parser();
            const string code = "Dim foo$\n";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var declaration = result.ChildNodes.FirstOrDefault() as DeclarationNode;
            Assert.IsNotNull(declaration);

            var identifier = declaration.ChildNodes.FirstOrDefault() as IdentifierNode;
            Assert.IsNotNull(identifier);

            Assert.AreEqual("foo", identifier.Name);
            Assert.AreEqual("String", identifier.TypeName);
            Assert.IsTrue(identifier.IsTypeSpecified);
        }

        [TestMethod]
        public void TestInitializerDeclaration()
        {
            var parser = new Parser();
            const string code = "Dim foo As New Collection\n";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var declaration = result.ChildNodes.FirstOrDefault() as DeclarationNode;
            Assert.IsNotNull(declaration);

            var identifier = declaration.ChildNodes.FirstOrDefault() as IdentifierNode;
            Assert.IsNotNull(identifier);

            Assert.AreEqual("foo", identifier.Name);
            Assert.AreEqual("Collection", identifier.TypeName);
            Assert.IsTrue(identifier.IsTypeSpecified);
            Assert.IsTrue(identifier.IsInitialized);
        }

        [TestMethod]
        public void TestLibraryReferenceDeclaration()
        {
            var parser = new Parser();
            const string code = "Dim foo As ADODB.Recordset\n";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var declaration = result.ChildNodes.FirstOrDefault() as DeclarationNode;
            Assert.IsNotNull(declaration);

            var identifier = declaration.ChildNodes.FirstOrDefault() as IdentifierNode;
            Assert.IsNotNull(identifier);

            Assert.AreEqual("foo", identifier.Name);
            Assert.AreEqual("ADODB.Recordset", identifier.TypeName);
            Assert.AreEqual("Rubberduck.Parser", identifier.Scope);
            Assert.AreEqual("ADODB", identifier.Library);
            Assert.IsTrue(identifier.IsTypeSpecified);
        }

        [TestMethod]
        public void TestMultipleDeclarations()
        {
            var parser = new Parser();
            const string code = "Dim foo, bar As String\n";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var declaration = result.ChildNodes.FirstOrDefault() as DeclarationNode;
            Assert.IsNotNull(declaration);

            var identifiers = declaration.ChildNodes.Select(node => node as IdentifierNode).ToList();
            Assert.AreEqual(2, identifiers.Count);

            Assert.AreEqual("foo", identifiers[0].Name);
            Assert.AreEqual("bar", identifiers[1].Name);

            Assert.AreEqual("Variant", identifiers[0].TypeName);
            Assert.AreEqual("String", identifiers[1].TypeName);
        }

        [TestMethod]
        public void TestMultipleDeclarationsWithArray()
        {
            var parser = new Parser();
            const string code = "Dim foo() As Integer, bar As String\n";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var declaration = result.ChildNodes.FirstOrDefault() as DeclarationNode;
            Assert.IsNotNull(declaration);

            var identifiers = declaration.ChildNodes.Select(node => node as IdentifierNode).ToList();
            Assert.AreEqual(2, identifiers.Count);

            Assert.AreEqual("foo", identifiers[0].Name);
            Assert.AreEqual("bar", identifiers[1].Name);

            Assert.IsTrue(identifiers[0].IsArray);
            Assert.IsFalse(identifiers[1].IsArray);
        }
        [TestMethod]
        public void TestConstDeclaration()
        {
            var parser = new Parser();
            const string code = "Const foo As String = \"test\"\n";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var declaration = result.ChildNodes.FirstOrDefault() as DeclarationNode;
            Assert.IsNotNull(declaration);

            var identifiers = declaration.ChildNodes.Select(node => node as IdentifierNode).ToList();

            Assert.AreEqual("foo", identifiers[0].Name);
            Assert.AreEqual("String", identifiers[0].TypeName);
        }

        [TestMethod]
        public void TestPublicSubIsProcedureNode()
        {
            var parser = new Parser();
            const string code = "Public Sub Foo()\n\rEnd Sub\n\r";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var procedure = result.ChildNodes.FirstOrDefault() as ProcedureNode;
            Assert.IsNotNull(procedure);
        }

        [TestMethod]
        public void TestProcedureNodeTakesParameter()
        {
            var parser = new Parser();
            const string code = "Public Sub Foo(bar)\n\rEnd Sub\n\r";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var procedure = (ProcedureNode)result.ChildNodes.First();
            Assert.AreEqual(1, procedure.Parameters.Count());
        }

        [TestMethod]
        public void TestProcedureNodeTakesParameters()
        {
            var parser = new Parser();
            const string code = "Public Sub Foo(ByVal a As Integer, ByRef b As Integer)\n\rEnd Sub\n\r";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var procedure = (ProcedureNode)result.ChildNodes.First();
            var parameters = procedure.Parameters.ToList();
            Assert.AreEqual(2, parameters.Count);
            Assert.AreEqual("a", parameters[0].Identifier.Name);
            Assert.AreEqual("b", parameters[1].Identifier.Name);
        }

        [TestMethod]
        public void ProcedureNodeHasChildren()
        {
            var parser = new Parser();
            const string code = "Public Sub Foo()\n\r    Dim bar As String\n\rEnd Sub\n\r";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var procedure = result.ChildNodes.First();
            Assert.IsTrue(procedure.ChildNodes.OfType<VariableDeclarationNode>().Count() == 1);
        }

        [TestMethod]
        public void ModuleHasProcedures()
        {
            var parser = new Parser();
            const string code = "Public Sub Foo()\n\r    Dim bar As String\n\rEnd Sub\n\rPublic Function Bar()\n\r    Dim foo As String\n\rEnd Function\n\r";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);

            var procedures = result.ChildNodes.OfType<ProcedureNode>().ToList();
            Assert.AreEqual(2, procedures.Count);
        }

        [TestMethod]
        public void InstructionHasIndentation()
        {
            var parser = new Parser();
            const string code = "Public Sub Foo()\n\r    Dim bar As String\n\rEnd Sub";

            var result = parser.Parse("ParserTests", "Rubberduck.Parser", code, false);
            var declaration = result.ChildNodes.OfType<ProcedureNode>().First()
                                    .ChildNodes.OfType<DeclarationNode>().First();

            Assert.AreEqual(5, declaration.Instruction.StartColumn);
        }
    }
}
