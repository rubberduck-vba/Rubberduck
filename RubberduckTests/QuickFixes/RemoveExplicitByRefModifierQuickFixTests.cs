﻿using System;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveExplicitByRefModifierQuickFixTests : QuickFixTestBase
    {

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_OptionalParameter()
        {
            const string inputCode =
                @"Sub Foo(Optional ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
                @"Sub Foo(Optional arg1 As Integer)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new RedundantByRefModifierInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_Optional_LineContinuations()
        {
            const string inputCode =
                @"Sub Foo(Optional ByRef _
        bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
                @"Sub Foo(Optional _
        bar _
        As Byte)
    bar = 1
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new RedundantByRefModifierInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_LineContinuation()
        {
            const string inputCode =
                @"Sub Foo( ByRef bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
                @"Sub Foo( bar _
        As Byte)
    bar = 1
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new RedundantByRefModifierInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_LineContinuation_FirstLine()
        {
            const string inputCode =
                @"Sub Foo( _
        ByRef bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
                @"Sub Foo( _
        bar _
        As Byte)
    bar = 1
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new RedundantByRefModifierInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_InterfaceImplementation()
        {
            const string interfaceInputCode =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string implementationInputCode =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedInterfaceCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string expectedImplementationCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new RedundantByRefModifierInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_InterfaceImplementationDiffrentParameterName()
        {
            const string interfaceInputCode =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string implementationInputCode =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg2 As Integer)
End Sub";

            const string expectedInterfaceCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string expectedImplementationCode =
                @"Implements IClass1

Sub IClass1_Foo(arg2 As Integer)
End Sub";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new RedundantByRefModifierInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_InterfaceImplementationWithMultipleParameters()
        {
            const string interfaceInputCode =
                @"Sub Foo(ByRef arg1 As Integer, ByRef arg2 as Integer)
End Sub";

            const string implementationInputCode =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer, ByRef arg2 as Integer)
End Sub";

            const string expectedInterfaceCode =
                @"Sub Foo(arg1 As Integer, ByRef arg2 as Integer)
End Sub";

            const string expectedImplementationCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer, ByRef arg2 as Integer)
End Sub";

            Func<IInspectionResult, bool> conditionOnResultToFix = result =>
                ((VBAParser.ArgContext)result.Context).unrestrictedIdentifier()
                .identifier()
                .untypedIdentifier()
                .identifierValue()
                .GetText() == "arg1";
            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterfaceSatisfyingPredicate(interfaceInputCode, implementationInputCode,
                    state => new RedundantByRefModifierInspection(state), conditionOnResultToFix);
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_PassByRef()
        {
            const string inputCode =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new RedundantByRefModifierInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveExplicitByRefModifierQuickFix(state);
        }
    }
}
