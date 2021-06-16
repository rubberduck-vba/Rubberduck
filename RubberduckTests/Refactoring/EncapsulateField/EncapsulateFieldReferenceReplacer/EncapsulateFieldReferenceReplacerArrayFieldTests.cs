﻿using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    /// <summary>
    /// EncapsulateFieldReferenceReplacerArrayFieldTests exclusively tests the reference update aspect
    /// of the EncapsulateFieldRefactoring.  So, refactored code generated by the tests do not include
    /// the encapsulation properties or modification of the target Declaration.  Only the target's IdentifierReferences 
    /// are updated (and checked).
    /// </summary>
    [TestFixture]
    public class EncapsulateFieldReferenceReplacerArrayFieldTests
    {
        private const bool IsReadOnlyEncapsulation = true;

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void PublicField_WrapAsPrivateUDT_DoesNotUsePropertiesWithinDeclaringModule()
        {
            var wrapInPrivateUDT = true;

            var testTargetTuple = ("mArray", "MyUDTMember", IsReadOnlyEncapsulation);

            var testModuleCode =
$@"
Option Explicit

Public mArray() As Integer

Private Sub InitializeArray(size As Long)
    Redim mArray(size)
    Dim idx As Long
    For idx = 1 To size
        mArray(idx) = idx
    Next idx
End Sub
";
            var declaringModule = ModuleTuple(MockVbeBuilder.TestModuleName, testModuleCode);

            var referencingModuleCode =
$@"
Option Explicit

Public Sub Fazz()
    Fazzle mArray(1)
End Sub

Public Sub Fazzle(arg As Integer)
End Sub
";
            var referencingModuleStdModule = ModuleTuple(moduleName: "SomeOtherModule", referencingModuleCode);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule, referencingModuleStdModule);

            var testModuleResult = refactoredCode[declaringModule.TestModuleName];
            StringAssert.Contains($"Redim this.MyUDTMember(size)", testModuleResult);
            StringAssert.Contains($"  this.MyUDTMember(idx) = idx", testModuleResult);

            var refModuleResult = refactoredCode[referencingModuleStdModule.TestModuleName];

            //Note: The refactoring forces the UDTMember Property identifier and UDTMember identifier to be the same
            StringAssert.Contains($"  Fazzle {declaringModule.TestModuleName}.MyUDTMember(1)", refModuleResult);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void PublicField_UseBackingField_DoesNotUsePropertiesWithinDeclaringModule()
        {
            var wrapInPrivateUDT = false;

            var testTargetTuple = ("mArray", "MyProperty", IsReadOnlyEncapsulation);

            var testModuleCode =
$@"
Option Explicit

Public mArray() As Integer

Private Sub InitializeArray(size As Long)
    Redim mArray(size)
    Dim idx As Long
    For idx = 1 To size
        mArray(idx) = idx
    Next idx
End Sub
";
            var declaringModule = ModuleTuple(MockVbeBuilder.TestModuleName, testModuleCode);

            var referencingModuleCode =
$@"
Option Explicit

Public Sub Fazz()
    Fazzle mArray(1)
End Sub

Public Sub Fazzle(arg As Integer)
End Sub
";
            var referencingModuleStdModule = ModuleTuple(moduleName: "SomeOtherModule", referencingModuleCode);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule, referencingModuleStdModule);

            var testModuleResult = refactoredCode[declaringModule.TestModuleName];
            StringAssert.Contains($"Redim mArray(size)", testModuleResult);
            StringAssert.Contains($"  mArray(idx) = idx", testModuleResult);

            var refModuleResult = refactoredCode[referencingModuleStdModule.TestModuleName];
            StringAssert.Contains($"  Fazzle {declaringModule.TestModuleName}.MyProperty(1)", refModuleResult);
        }

        private static (string TestModuleName, string TargetID, ComponentType ComponentType) ModuleTuple(string moduleName, string targetName, ComponentType componentType = ComponentType.StandardModule)
            => (moduleName, targetName, componentType);

        private static IDictionary<string, string> TestReferenceReplacement(bool wrapInPrivateUDT, (string, string, bool) testTargetTuple, params (string, string, ComponentType)[] moduleTuples )
        {
            return ReferenceReplacerTestSupport.TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, moduleTuples);
        }
    }
}