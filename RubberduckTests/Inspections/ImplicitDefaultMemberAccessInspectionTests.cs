﻿using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplicitDefaultMemberAccessInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        [TestCase("Foo = cls")]
        [TestCase("cls = bar")]
        public void OrdinaryImplicitDefaultMemberAccessExpression_OneResult(string statement)
        {
            var class1Code = @"
Public Function Foo() As Long
Attribute Foo.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function Foo() As Long 
    Dim cls As New Class1
    Dim bar As Long
    {statement}
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Foo = cls")]
        [TestCase("cls = bar")]
        //There is a specialized inspection for this.
        public void UnboundImplicitDefaultMemberAccessExpression_NoResult(string statement)
        {
            var class1Code = @"
Public Function Foo() As Long
Attribute Foo.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function Foo() As Long 
    Dim cls As Object
    Set cls = New Class1
    Dim bar As Long
    {statement}
End Function
";


            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        //There is a specialized inspection for this.
        public void ProcedureCoercion_NoResult()
        {
            var class1Code = @"
Public Function Foo() As Long
Attribute Foo.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function Foo() As Long 
    Dim cls As New Class1
    cls
End Function
";


            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Foo = cls")]
        [TestCase("cls = bar")]
        //There is a specialized inspection for this.
        public void RecursiveImplicitDefaultMemberAccessExpression_NoResult(string statement)
        {
            var class1Code = @"
Public Function Foo() As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz() As Long
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function Foo() As Long 
    Dim cls As New Class1
    Dim bar As Long
    {statement}
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Foo = cls(0)")]
        [TestCase("cls(0) = bar")]
        public void OrdinaryImplicitDefaultMemberAccessOnDefaultMemberArrayAccess_TwoResults(string statement)
        {
            var class1Code = @"
Public Function Foo() As Class2()
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz() As Long
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function Foo() As Long 
    Dim cls As New Class1
    Dim bar As Long
    {statement}
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var inspectionResults = InspectionResults(vbe.Object);

            Assert.AreEqual(2, inspectionResults.Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ImplicitDefaultMemberAccessInspection(state);
        }
    }
}