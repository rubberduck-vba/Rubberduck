using System;
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
    [TestFixture]
    public class EncapsulateFieldReferenceReplacerArrayFieldTests
    {
        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void PublicField_DoesNotUsePropertiesWithinDeclaringModule(bool wrapInPrivateUDT)
        {
            var target = "mArray";
            var propertyName = "MyProperty";

            var testTargetTuple = (target, propertyName, true);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
        Option Explicit

        Public {target}() As Integer

        Private Sub InitializeArray(size As Long)
            Redim {target}(size)
            Dim idx As Long
            For idx = 1 To size
                {target}(idx) = idx
            Next idx
        End Sub
        ";
            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var referencingModule = "SomeOtherModule";
            var referencingModuleCode =
$@"
        Option Explicit

        Public Sub Fazz()
            Fazzle {target}(1)
        End Sub

        Public Sub Fazzle(arg As Integer)
        End Sub
        ";

            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);

            var refactoredCode = TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, declaringModule, referencingModuleStdModule);

            var testModuleResult = refactoredCode[testModuleName];
            var refModuleResult = refactoredCode[referencingModule];

            if (wrapInPrivateUDT)
            {
                StringAssert.Contains($"  Fazzle {testModuleName}.{propertyName}(1)", refModuleResult);

                StringAssert.Contains($"Redim this.{propertyName}(size)", testModuleResult);
                StringAssert.Contains($"  this.{propertyName}(idx) = idx", testModuleResult);
                return;
            }

            StringAssert.Contains($"  Fazzle {testModuleName}.{propertyName}(1)", refModuleResult);

            StringAssert.Contains($"Redim {target}(size)", testModuleResult);
            StringAssert.Contains($"  {target}(idx) = idx", testModuleResult);
        }

        private static IDictionary<string, string> TestReferenceReplacement(bool wrapInPrivateUDT, (string, string, bool) testTargetTuple, params (string, string, ComponentType)[] moduleTuples )
        {
            return ReferenceReplacerTestSupport.TestReferenceReplacement(wrapInPrivateUDT, testTargetTuple, moduleTuples);
        }
    }
}
