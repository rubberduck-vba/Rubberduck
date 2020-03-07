using System;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class SpecifyExplicitByRefModifierQuickFixTests : QuickFixTestBase
    {

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_ImplicitByRefParameter()
        {
            const string inputCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string expectedCode =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ImplicitByRefModifierInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_OptionalParameter()
        {
            const string inputCode =
                @"Sub Foo(Optional arg1 As Integer)
End Sub";

            const string expectedCode =
                @"Sub Foo(Optional ByRef arg1 As Integer)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ImplicitByRefModifierInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_Optional_LineContinuations()
        {
            const string inputCode =
                @"Sub Foo(Optional _
        bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
                @"Sub Foo(Optional _
        ByRef bar _
        As Byte)
    bar = 1
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ImplicitByRefModifierInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_LineContinuation()
        {
            const string inputCode =
                @"Sub Foo(bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
                @"Sub Foo(ByRef bar _
        As Byte)
    bar = 1
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ImplicitByRefModifierInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_LineContinuation_FirstLine()
        {
            const string inputCode =
                @"Sub Foo( _
        bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
                @"Sub Foo( _
        ByRef bar _
        As Byte)
    bar = 1
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ImplicitByRefModifierInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_InterfaceImplementation()
        {
            const string interfaceInputCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string implementationInputCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            const string expectedInterfaceCode =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedImplementationCode =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new ImplicitByRefModifierInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_InterfaceImplementationDifferentParameterName()
        {
            const string interfaceInputCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string implementationInputCode =
                @"Implements IClass1

Sub IClass1_Foo(arg2 As Integer)
End Sub";

            const string expectedInterfaceCode =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedImplementationCode =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg2 As Integer)
End Sub";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new ImplicitByRefModifierInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_InterfaceImplementationWithMultipleParameters()
        {
            const string interfaceInputCode =
                @"Sub Foo(arg1 As Integer, arg2 as Integer)
End Sub";

            const string implementationInputCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer, arg2 as Integer)
End Sub";

            const string expectedInterfaceCode =
                @"Sub Foo(ByRef arg1 As Integer, arg2 as Integer)
End Sub";

            const string expectedImplementationCode =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer, arg2 as Integer)
End Sub";

            Func<IInspectionResult,bool> conditionOnResultToFix = result =>
                ((VBAParser.ArgContext)result.Context).unrestrictedIdentifier()
                .identifier()
                .untypedIdentifier()
                .identifierValue()
                .GetText() == "arg1";
            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterfaceSatisfyingPredicate(interfaceInputCode, implementationInputCode,
                    state => new ImplicitByRefModifierInspection(state), conditionOnResultToFix);
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new SpecifyExplicitByRefModifierQuickFix(state);
        }
    }
}
