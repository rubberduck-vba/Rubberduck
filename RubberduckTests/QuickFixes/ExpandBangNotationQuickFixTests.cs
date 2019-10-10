﻿using NUnit.Framework;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    public class ExpandBangNotationQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void DictionaryAccessExpression_QuickFixWorks()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls!newClassObject
End Function
";

            var expectedModuleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls.Foo(""newClassObject"")
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new UseOfBangNotationInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RecursiveDictionaryAccessExpression_QuickFixWorks()
        {
            var class1Code = @"
Public Function Foo() As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls!newClassObject
End Function
";

            var expectedModuleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls.Foo().Baz(""newClassObject"")
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new UseOfRecursiveBangNotationInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void WithDictionaryAccessExpression_QuickFixWorks()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    With New Class1
        Set Foo = !newClassObject
    End With
End Function
";

            var expectedModuleCode = @"
Private Function Foo() As Class2 
    With New Class1
        Set Foo = .Foo(""newClassObject"")
    End With
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new UseOfBangNotationInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RecursiveWithDictionaryAccessExpression_QuickFixWorks()
        {
            var class1Code = @"
Public Function Foo() As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    With New Class1
        Set Foo = !newClassObject
    End With
End Function
";

            var expectedModuleCode = @"
Private Function Foo() As Class2 
    With New Class1
        Set Foo = .Foo().Baz(""newClassObject"")
    End With
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new UseOfRecursiveBangNotationInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ExpandBangNotationQuickFix(state);
        }
    }
}