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
    public class EncapsulateFieldReferenceReplacerTests
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PublicValueField_ExternalReference()
        {
            var target = "targetField";
            var propertyName = "MyProperty";

            var testModuleName = MockVbeBuilder.TestModuleName;
            var referenceExpression = $"{testModuleName}.{target}";
            var testModuleCode =
$@"
Option Explicit
Public targetField As Long";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);


            var procedureModuleReferencingCode =
$@"
Option Explicit

Public Sub Bar()
    {referenceExpression} = 7
End Sub
";
            var referencingModuleStdModule = (moduleName: "StdModule", procedureModuleReferencingCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var refactoredCode = ReplaceReferences(vbe.Object, (target, propertyName, false));

            var referencingModuleCode = refactoredCode[referencingModuleStdModule.moduleName];

            StringAssert.Contains($"{testModuleName}.{propertyName} = ", referencingModuleCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PublicValueField_ExternalWithMemberAccessReference()
        {
            var target = "targetField";
            var propertyName = "MyProperty";

            var testModuleName = MockVbeBuilder.TestModuleName;
            var referenceExpression = $".{target}";
            var testModuleCode =
$@"
Option Explicit
Public targetField As Long";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);


            var procedureModuleReferencingCode =
$@"
Option Explicit

Public Sub Bar()
    With {testModuleName}
        {referenceExpression} = 7
    End With
End Sub
";
            var referencingModuleStdModule = (moduleName: "StdModule", procedureModuleReferencingCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var refactoredCode = ReplaceReferences(vbe.Object, (target, propertyName, false));

            var referencingModuleCode = refactoredCode[referencingModuleStdModule.moduleName];

            StringAssert.Contains($"  .{propertyName} = ", referencingModuleCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void ValueField_LocalReference(string visibility)
        {
            var target = "targetField";
            var propertyName = "MyProperty";

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Option Explicit
{visibility} targetField As Long

Public Sub Bar()
    {target} = 7
    Bars {target}
End Sub

Public Sub Bars(ByVal arg As Long)
End Sub
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule);
            var refactoredCode = ReplaceReferences(vbe.Object, (target, propertyName, false));

            var referencingModuleCode = refactoredCode[testModuleName];

            StringAssert.Contains($"{propertyName} = ", referencingModuleCode);
            StringAssert.Contains($" Bars {propertyName}", referencingModuleCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void ValueField_LocalReferenceReadOnly(string visibility)
        {
            var target = "targetField";
            var propertyName = "MyProperty";

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Option Explicit
{visibility} targetField As Long

Public Sub Bar()
    {target} = 7
    Bars {target}
End Sub

Public Sub Bars(ByVal arg As Long)
End Sub
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule);
            var refactoredCode = ReplaceReferences(vbe.Object, (target, propertyName, true));

            var referencingModuleCode = refactoredCode[testModuleName];

            StringAssert.Contains($"{target} = ", referencingModuleCode);
            StringAssert.Contains($" Bars {propertyName}", referencingModuleCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PublicUDTField_ExternalReference()
        {
            var target = "targetField";
            var propertyName = "MyProperty";

            var testModuleName = MockVbeBuilder.TestModuleName;
            var referenceExpression = $"{testModuleName}.{target}";
            var testModuleCode =
$@"
Option Explicit

Public Type TestType
    Fizz As Long
End Type

Public targetField As TestType";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);


            var procedureModuleReferencingCode =
$@"
Option Explicit

Public Sub Bar()
    {referenceExpression}.Fizz = 7
End Sub
";
            var referencingModuleStdModule = (moduleName: "StdModule", procedureModuleReferencingCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var refactoredCode = ReplaceReferences(vbe.Object, (target, propertyName, false));

            var referencingModuleCode = refactoredCode[referencingModuleStdModule.moduleName];

            StringAssert.Contains($"{testModuleName}.{propertyName}.Fizz = ", referencingModuleCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTField_PublicType_StdModuleReferenceWithMemberAccess()
        {
            var target = "targetField";
            var propertyName = "MyProperty";

            var testModuleName = MockVbeBuilder.TestModuleName;

            var testModuleCode =
$@"
Public Type TBar
    First As String
    Second As Long
End Type

Public targetField As TBar";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var moduleReferencingCode =
$@"Option Explicit

'StdModule referencing the UDT

Public Sub FooBar()
    With {testModuleName}
        .targetField.First = ""Foo""
        .targetField.Second = 7
    End With
End Sub
";
            var referencingModuleStdModule = (moduleName: "StdModule", moduleReferencingCode, ComponentType.StandardModule);
            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var refactoredCode = ReplaceReferences(vbe.Object, (target, propertyName, false));

            var referencingModuleCode = refactoredCode[referencingModuleStdModule.moduleName];

            StringAssert.Contains($"  .MyProperty.First = ", referencingModuleCode);
            StringAssert.Contains($"  .MyProperty.Second = ", referencingModuleCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTFieldSelection_ClassModuleSource_ExternalReferences()
        {
            var target = "targetField";
            var propertyName = "MyProperty";

            var testModuleName = MockVbeBuilder.TestModuleName;
            var classInstanceName = "theClass";
            var testModuleCode =
$@"
Option Explicit

Public targetField As TBar";
            var declaringModule = (testModuleName, testModuleCode, ComponentType.ClassModule);

            var moduleReferencingCode =
$@"
Option Explicit

Public Type TBar
    First As String
    Second As Long
End Type

Private {classInstanceName} As {testModuleName}

Public Sub Initialize()
    Set {classInstanceName} = New {testModuleName}
End Sub

Public Sub Fizz()
    {classInstanceName}.targetField.First = ""Foo""
End Sub

Public Sub Bang()
    {classInstanceName}.targetField.Second = 7
End Sub

Public Sub FizzBang()
    With {classInstanceName}
        .targetField.First = ""FizzBang""
        .targetField.Second = 7
    End With
End Sub
";

            var referencingModuleStdModule = (moduleName: "StdModule", moduleReferencingCode, ComponentType.StandardModule);
            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var refactoredCode = ReplaceReferences(vbe.Object, (target, propertyName, false));

            var referencingModuleCode = refactoredCode[referencingModuleStdModule.moduleName];
            StringAssert.Contains($"{classInstanceName}.{propertyName}.First = ", referencingModuleCode);
            StringAssert.Contains($"{classInstanceName}.{propertyName}.Second = ", referencingModuleCode);
            StringAssert.Contains($"  .{propertyName}.Second = ", referencingModuleCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void OverrideReadOnlyFlagForExternalReferences()
        {
            var target = "targetField";
            var propertyName = "MyProperty";

            //Simulates the scenario where the readOnly flag was set to 'True' by some means (other than the UI).
            //Since there are external references, a property Let/Set will be generated 
            //by the refactoring - so the references are modified accordingly
            var testTargetTuple = (target, propertyName, true);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Option Explicit

Public targetField As Long

Private mValue As Long

Public Sub Fizz(arg As Long)
    mValue = targetField + arg
End Sub
";
            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var referencingModule = "SomeOtherModule";
            var referencingModuleCode =
$@"
Option Explicit

Private mValue As Long

Public Sub Fazz(arg As Long)
    mValue = targetField * arg
End Sub

Public Sub Fazzle(arg As Long)
    targetField = arg * mValue
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);
            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);

            StringAssert.Contains($"mValue = {propertyName} + arg", refactoredCode[MockVbeBuilder.TestModuleName]);

            StringAssert.Contains($"mValue = {MockVbeBuilder.TestModuleName}.{propertyName} * arg", refactoredCode[referencingModule]);
            StringAssert.Contains($"  {MockVbeBuilder.TestModuleName}.{propertyName} = arg * mValue", refactoredCode[referencingModule]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void ArrayReferences()
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
            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);

            var testModuleResult = refactoredCode[testModuleName];
            var refModuleResult = refactoredCode[referencingModule];

            StringAssert.Contains($"  Fazzle {testModuleName}.{propertyName}(1)", refModuleResult);

            StringAssert.Contains($"Redim {target}(size)", testModuleResult);
            StringAssert.Contains($"  {target}(idx) = idx", testModuleResult);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void ReplaceUdtMemberReferences()
        {
            var target = "myBazz";
            var propertyName = "MyProperty";

            var testTargetTuple = (target, propertyName, false);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Private Type TBazz
    FirstValue As String
    SecondValue As Long
End Type

Private myBazz As TBazz

Public Sub Fizz(newValue As String)
    myBazz.FirstValue = newValue
End Sub

Public Sub Bazz(newValue As Long)
    myBazz.SecondValue = newValue
End Sub
";
            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule);
            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);

            StringAssert.Contains($"  FirstValue = newValue", refactoredCode[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains($"  SecondValue = newValue", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void RenameFieldReferences_WithMemberAccess_NoExternalReferences(bool isReadOnly)
        {
            var target = "myBazz";
            var propertyName = "MyProperty";

            var testTargetTuple = (target, propertyName, isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Private Type TBazz
    FirstValue As String
    SecondValue As Long
End Type

Private myBazz As TBazz

Public Sub Fizz(newValue As String)
    With myBazz
        .FirstValue = newValue
    End With
End Sub

Public Sub Bazz(newValue As String)
    With myBazz
        .SecondValue = newValue
    End With
End Sub
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule);
            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);

            StringAssert.Contains($"  With myBazz{Environment.NewLine}", refactoredCode[MockVbeBuilder.TestModuleName]);
            if (isReadOnly) //Get generated
            {
                StringAssert.Contains("  .FirstValue = newValue", refactoredCode[MockVbeBuilder.TestModuleName]);
                StringAssert.Contains("  .SecondValue = newValue", refactoredCode[MockVbeBuilder.TestModuleName]);
            }
            else //Let and Get generated
            {
                //The EF refactoring will create a FirstValue and SecondValue property - so the with member access expression
                //is replaced with the Let property name. The EF refactoring does not remove the 'With' statement block even 
                //though it is no longer required by this specific scenario
                StringAssert.Contains("  FirstValue = newValue", refactoredCode[MockVbeBuilder.TestModuleName]);
                StringAssert.Contains("  SecondValue = newValue", refactoredCode[MockVbeBuilder.TestModuleName]);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void ReplaceAccessorExpression()
        {
            var target = "myBazz";
            var propertyName = "MyProperty";

            var testTargetTuple = (target, propertyName, false);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Private Type TBazz
    FirstValue As String
    SecondValue As Long
End Type

Private myBazz As TBazz

Public Function GetTheFirstValue() As String
    GetTheFirstValue = myBazz.FirstValue
End Function

Public Sub SetTheFirstValue(arg As Long)
    myBazz.FirstValue = arg
End Sub
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule);
            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);

            StringAssert.Contains("FirstValue = arg", refactoredCode[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains("GetTheFirstValue = FirstValue", refactoredCode[MockVbeBuilder.TestModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void ReplacePublicUDTAccessorExpression()
        {
            var target = "myBazz";
            var propertyName = "MyProperty";

            var testTargetTuple = (target, propertyName, false);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"

Public Type TBazz
    FirstValue As String
    SecondValue As Long
End Type

Public myBazz As TBazz

Public Function GetTheFirstValue() As String
    GetTheFirstValue = myBazz.FirstValue
End Function

Public Sub SetTheFirstValue(arg As Long)
    myBazz.FirstValue = arg
End Sub
";
            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var referencingModule = "AnotherModule";
            var referencingModuleCode =
$@"

Public Function GetBazzFirst() As String
    GetBazzFirst = myBazz.FirstValue
End Function

Public Sub SetBazzFirst(arg As Long)
    myBazz.FirstValue = arg
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);

            StringAssert.Contains($"{propertyName}.FirstValue = arg", refactoredCode[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains($"GetTheFirstValue = {propertyName}.FirstValue", refactoredCode[MockVbeBuilder.TestModuleName]);
            StringAssert.Contains($"GetBazzFirst = {MockVbeBuilder.TestModuleName}.{propertyName}.FirstValue", refactoredCode[referencingModule]);
            StringAssert.Contains($"{MockVbeBuilder.TestModuleName}.{propertyName}.FirstValue = arg", refactoredCode[referencingModule]);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void ModifiesCorrectUDTMemberReferences_MemberAccess(bool isReadOnly)
        {
            var target = "this";

            var testTargetTuple = (target, string.Empty, isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private this As TBar

Private that As TBar

Public Sub Foo(arg1 As String, arg2 As Long)
    this.First = arg1
    that.First = arg1
    this.Second = arg2
    that.Second = arg2
End Sub
";
            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule);
            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);

            var actualCode = refactoredCode[testModuleName];
            if (isReadOnly)
            {
                StringAssert.Contains($" this.First = arg1", actualCode);
                StringAssert.Contains($" this.Second = arg2", actualCode);
            }
            else
            {
                StringAssert.Contains($" First = arg1", actualCode);
                StringAssert.Contains($" Second = arg2", actualCode);
            }
            StringAssert.Contains($"that.First = arg1", actualCode);
            StringAssert.Contains($"that.Second = arg2", actualCode);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void NestedUDTMember(bool isReadOnly)
        {
            var target = "mTypesField";

            var testTargetTuple = (target, string.Empty, isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Option Explicit

Private Type PType1
    FirstValType1 As Long
    SecondValType1 As String
End Type

Private Type PType2
    FirstValType2 As Long
    SecondValType2 As String
    Third As PType1
End Type

Private mTypesField As PType2

Private Sub Class_Initialize()
    mTypesField.Third.SecondValType1 = ""Wah""
End Sub

Private Sub TestSub2()
    TestSub3 mTypesField.Third.SecondValType1
End Sub

Private Sub TestSub3(ByVal arg As String)
End Sub
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule);
            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);

            var expectedAssignment = isReadOnly ? "TypesField.Third.SecondValType1 = \"Wah\"" : "SecondValType1 = \"Wah\"";
            StringAssert.Contains(expectedAssignment, refactoredCode[testModuleName]);

            StringAssert.Contains("TestSub3 SecondValType1", refactoredCode[testModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void PublicUDTField_ExternalRefNestedWithStatements()
        {
            var target = "mTypesField";

            var testTargetTuple = (target, "TypesField", false);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Option Explicit

Public Type PType1
    FirstValType1 As Long
End Type

Public Type PType2
    Third As PType1
End Type

Public mTypesField As PType2
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var referencingModuleName = "AnotherModule";
            var referencingModuleCode =
$@"
Private testVal As Long

Public Sub TestSub()
    With {testModuleName}
        With .mTypesField
            With .Third
                testVal = .FirstValType1
            End With
        End With
    End With
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModuleName, referencingModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);

            StringAssert.Contains($"With .TypesField", refactoredCode[referencingModuleName]);

            StringAssert.Contains(" testVal = .FirstValType1", refactoredCode[referencingModuleName]);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldReferenceReplacer))]
        public void PrivateUDTField_RefNestedWithStatements(bool isReadOnly)
        {
            var target = "mTypesField";

            var testTargetTuple = (target, "TypesField", isReadOnly);

            var testModuleName = MockVbeBuilder.TestModuleName;
            var testModuleCode =
$@"
Option Explicit

Private Type PType1
    FirstValType1 As Long
End Type

Private Type PType2
    Third As PType1
End Type

Private mTypesField As PType2

Public Sub TestSub(ByVal arg As Long)
    With mTypesField
        With .Third
            .FirstValType1 = arg
        End With
    End With
End Sub

Public Function TestFunc() As Long
    With mTypesField
        With .Third
            TestFunc = .FirstValType1
        End With
    End With
End Function

";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule);
            var refactoredCode = ReplaceReferences(vbe.Object, testTargetTuple);

            StringAssert.Contains($"With mTypesField", refactoredCode[testModuleName]);
            StringAssert.Contains($"With .Third", refactoredCode[testModuleName]);

            var expectedAssignment = isReadOnly ? ".FirstValType1 = arg" : "FirstValType1 = arg";
            StringAssert.Contains(expectedAssignment, refactoredCode[testModuleName]);

            StringAssert.Contains("TestFunc = FirstValType1", refactoredCode[testModuleName]);
        }

        private IDictionary<string, string> ReplaceReferences(IVBE vbe, (string fieldID, string fieldProperty, bool readOnly) target, params (string fieldID, string fieldProperty, bool readOnly)[] fieldIDPairs) 
            => ReplaceReferences(vbe, target, fieldIDPairs.ToList());
        private IDictionary<string, string> ReplaceReferences(IVBE vbe, (string fieldID, string fieldProperty, bool readOnly) target, IEnumerable<(string fieldID, string fieldProperty, bool readOnly)> fieldIDPairs)
        {
            var refactoredCode = new Dictionary<string, string>();
            (var state, var rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                var resolver = Support.SetupResolver(state, rewritingManager);
                var encapsulateFieldFactory = resolver.Resolve<IEncapsulateFieldCandidateFactory>();
                var sutFactory = resolver.Resolve<IEncapsulateFieldReferenceReplacerFactory>();

                var fieldCandidate = encapsulateFieldFactory.CreateFieldCandidate(state.DeclarationFinder.MatchName(target.fieldID).Single());
                fieldCandidate.PropertyIdentifier = target.fieldProperty;
                fieldCandidate.IsReadOnly = target.readOnly;
                fieldCandidate.EncapsulateFlag = true;

                var sut = sutFactory.Create();
                sut.ReplaceReferences(new IEncapsulateFieldCandidate[] { fieldCandidate }, rewriteSession);


                if (rewriteSession.TryRewrite())
                {
                    refactoredCode = vbe.ActiveVBProject.VBComponents
                        .ToDictionary(component => component.Name, component => component.CodeModule.Content());
                }
            }

            return refactoredCode;
        }
    }
}
