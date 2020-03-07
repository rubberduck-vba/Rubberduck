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
    public class ChangeIntegerToLongQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_Function()
        {
            const string inputCode =
                @"Function Foo() As Integer
End Function";

            const string expectedCode =
                @"Function Foo() As Long
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_FunctionWithTypeHint()
        {
            const string inputCode =
                @"Function Foo%()
End Function";

            const string expectedCode =
                @"Function Foo&()
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_PropertyGet()
        {
            const string inputCode =
                @"Property Get Foo() As Integer
End Property";

            const string expectedCode =
                @"Property Get Foo() As Long
End Property";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_PropertyGetWithTypeHint()
        {
            const string inputCode =
                @"Property Get Foo%()
End Property";

            const string expectedCode =
                @"Property Get Foo&()
End Property";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_Parameter()
        {
            const string inputCode =
                @"Sub Foo(arg As Integer)
End Sub";

            const string expectedCode =
                @"Sub Foo(arg As Long)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ParameterWithTypeHint()
        {
            const string inputCode =
                @"Sub Foo(arg%)
End Sub";

            const string expectedCode =
                @"Sub Foo(arg&)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_Variable()
        {
            const string inputCode =
                @"Sub Foo()
    Dim v As Integer
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim v As Long
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_VariableWithTypeHint()
        {
            const string inputCode =
                @"Sub Foo()
    Dim v%
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim v&
End Sub";
            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_Constant()
        {
            const string inputCode =
                @"Sub Foo()
    Const c As Integer = 0
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Const c As Long = 0
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ConstantWithTypeHint()
        {
            const string inputCode =
                @"Sub Foo()
    Const c% = 0
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Const c& = 0
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_UserDefinedTypeReservedNameMember()
        {
            const string inputCode =
                @"Type T
    i as Integer
End Type";

            const string expectedCode =
                @"Type T
    i as Long
End Type";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_UserDefinedTypeUntypedNameMember()
        {
            const string inputCode =
                @"Type T
    i() as Integer
End Type";

            const string expectedCode =
                @"Type T
    i() as Long
End Type";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_FunctionInterfaceImplementation()
        {
            const string interfaceInputCode =
                @"Function Foo() As Integer
End Function";

            const string implementationInputCode =
                @"Implements IClass1

Function IClass1_Foo() As Integer
End Function";

            const string expectedInterfaceCode =
                @"Function Foo() As Long
End Function";

            const string expectedImplementationCode =
                @"Implements IClass1

Function IClass1_Foo() As Long
End Function";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_FunctionInterfaceImplementationWithTypeHints()
        {
            const string interfaceInputCode =
                @"Function Foo%()
End Function";

            const string implementationInputCode =
                @"Implements IClass1

Function IClass1_Foo%()
End Function";

            const string expectedInterfaceCode =
                @"Function Foo&()
End Function";

            const string expectedImplementationCode =
                @"Implements IClass1

Function IClass1_Foo&()
End Function";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_FunctionInterfaceImplementationWithInterfaceTypeHint()
        {
            const string interfaceInputCode =
                @"Function Foo%()
End Function";

            const string implementationInputCode =
                @"Implements IClass1

Function IClass1_Foo() As Integer
End Function";

            const string expectedInterfaceCode =
                @"Function Foo&()
End Function";

            const string expectedImplementationCode =
                @"Implements IClass1

Function IClass1_Foo() As Long
End Function";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_FunctionInterfaceImplementationWithImplementationTypeHint()
        {
            const string interfaceInputCode =
                @"Function Foo() As Integer
End Function";

            const string implementationInputCode =
                @"Implements IClass1

Function IClass1_Foo%()
End Function";

            const string expectedInterfaceCode =
                @"Function Foo() As Long
End Function";

            const string expectedImplementationCode =
                @"Implements IClass1

Function IClass1_Foo&()
End Function";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_PropertyGetInterfaceImplementation()
        {
            const string interfaceInputCode =
                @"Property Get Foo() As Integer
End Property";

            const string implementationInputCode =
                @"Implements IClass1

Property Get IClass1_Foo() As Integer
End Property";

            const string expectedInterfaceCode =
                @"Property Get Foo() As Long
End Property";

            const string expectedImplementationCode =
                @"Implements IClass1

Property Get IClass1_Foo() As Long
End Property";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_PropertyGetInterfaceImplementationWithTypeHints()
        {
            const string interfaceInputCode =
                @"Property Get Foo%()
End Property";

            const string implementationInputCode =
                @"Implements IClass1

Property Get IClass1_Foo%()
End Property";

            const string expectedInterfaceCode =
                @"Property Get Foo&()
End Property";

            const string expectedImplementationCode =
                @"Implements IClass1

Property Get IClass1_Foo&()
End Property";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_PropertyGetInterfaceImplementationWithInterfaceTypeHint()
        {
            const string interfaceInputCode =
                @"Property Get Foo%()
End Property";

            const string implementationInputCode =
                @"Implements IClass1

Property Get IClass1_Foo() As Integer
End Property";

            const string expectedInterfaceCode =
                @"Property Get Foo&()
End Property";

            const string expectedImplementationCode =
                @"Implements IClass1

Property Get IClass1_Foo() As Long
End Property";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_PropertyGetInterfaceImplementationWithImplementationTypeHint()
        {
            const string interfaceInputCode =
                @"Property Get Foo() As Integer
End Property";

            const string implementationInputCode =
                @"Implements IClass1

Property Get IClass1_Foo%()
End Property";

            const string expectedInterfaceCode =
                @"Property Get Foo() As Long
End Property";

            const string expectedImplementationCode =
                @"Implements IClass1

Property Get IClass1_Foo&()
End Property";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ParameterInterfaceImplementationWithTypeHints()
        {
            const string interfaceInputCode =
                @"Sub Foo(arg1%)
End Sub";

            const string implementationInputCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1%)
End Sub";

            const string expectedInterfaceCode =
                @"Sub Foo(arg1&)
End Sub";

            const string expectedImplementationCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1&)
End Sub";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ParameterInterfaceImplementationWithInterfaceTypeHint()
        {
            const string interfaceInputCode =
                @"Sub Foo(arg1%)
End Sub";

            const string implementationInputCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            const string expectedInterfaceCode =
                @"Sub Foo(arg1&)
End Sub";

            const string expectedImplementationCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Long)
End Sub";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ParameterInterfaceImplementationWithImplementationTypeHint()
        {
            const string interfaceInputCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string implementationInputCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1%)
End Sub";

            const string expectedInterfaceCode =
                @"Sub Foo(arg1 As Long)
End Sub";

            const string expectedImplementationCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1&)
End Sub";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ParameterInterfaceImplementation()
        {
            const string interfaceInputCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string implementationInputCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            const string expectedInterfaceCode =
                @"Sub Foo(arg1 As Long)
End Sub";

            const string expectedImplementationCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Long)
End Sub";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ParameterInterfaceImplementationWithDifferentName()
        {
            const string interfaceInputCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string implementationInputCode =
                @"Implements IClass1

Sub IClass1_Foo(arg2 As Integer)
End Sub";

            const string expectedInterfaceCode =
                @"Sub Foo(arg1 As Long)
End Sub";

            const string expectedImplementationCode =
                @"Implements IClass1

Sub IClass1_Foo(arg2 As Long)
End Sub";

            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterface(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state));
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_MultipleParameterInterfaceImplementation()
        {
            const string interfaceInputCode =
                @"Sub Foo(arg1 As Integer, arg2 as Integer)
End Sub";

            const string implementationInputCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer, arg2 as Integer)
End Sub";

            const string expectedInterfaceCode =
                @"Sub Foo(arg1 As Long, arg2 as Integer)
End Sub";

            const string expectedImplementationCode =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Long, arg2 as Integer)
End Sub";

            Func<IInspectionResult, bool> conditionOnResultToFix = result =>
                ((VBAParser.ArgContext)result.Context).unrestrictedIdentifier()
                .identifier()
                .untypedIdentifier()
                .identifierValue()
                .GetText() == "arg1";
            var (actualInterfaceCode, actualImplementationCode) =
                ApplyQuickFixToFirstInspectionResultForImplementedInterfaceSatisfyingPredicate(interfaceInputCode, implementationInputCode,
                    state => new IntegerDataTypeInspection(state), conditionOnResultToFix);
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedImplementationCode, actualImplementationCode, "Wrong code in implementation");
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ChangeIntegerToLongQuickFix(state);
        }
    }
}
