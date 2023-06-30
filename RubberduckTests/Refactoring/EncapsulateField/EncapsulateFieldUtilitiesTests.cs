using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateFieldUtilitiesTests
    {
        //Nested UDT declaration used by multiple tests
        private static string privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue =
$@"
Private Type AType
    AValue As Long
End Type

Private Type BType
    BValue As Long
    AInst As AType
End Type

Private Type CType
    CValue As Long
    BInst As BType
End Type

Private Type DType
    DValue As Long
    CInst As CType
End Type
";

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void GetCorrectReferenceCount()
        {
            var inputCode =
$@"
Private Type TBizz
    FirstVal As String
    SecondVal As Long
    ThirdVal As Byte
End Type

Private myBizz As TBizz
Private myFizz As TBizz

Public Function GetOne() As String
    GetOne = myBizz.FirstVal
End Function

Public Function GetTwo() As Long
    GetTwo = myBizz.ThirdVal
End Function
";
            var actual = GetReferenceCount(inputCode, "myBizz", "TBizz");
            Assert.AreEqual(2, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void GetCorrectReferenceCountPerInstance()
        {
            var inputCode =
$@"
Private Type TBizz
    FirstVal As String
    SecondVal As Long
    ThirdVal As Byte
End Type

Private myBizz As TBizz
Private myFizz As TBizz

Public Function GetOne() As String
    GetOne = myBizz.FirstVal
End Function

Public Function GetTwo() As Long
    GetTwo = myBizz.ThirdVal
End Function

Public Function GetThree() As Long
    GetThree = myFizz.ThirdVal
End Function

";

            var actual = GetReferenceCount(inputCode, "myBizz", "TBizz");
            Assert.AreEqual(2, actual);

            actual = GetReferenceCount(inputCode, "myFizz", "TBizz");
            Assert.AreEqual(1, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void GetCorrectReferenceCount_WithMemberAccess()
        {
            var inputCode =
$@"
Private Type TBizz
    FirstVal As String
    SecondVal As Long
    ThirdVal As Byte
End Type

Private myBizz As TBizz
Private myFizz As TBizz

Public Function GetOne() As String
    With myBizz
        GetOne = .FirstVal
    End With
End Function

Public Function GetTwo() As Long
    With myBizz
        GetTwo = .SecondVal
    End With
End Function

Public Function GetThree() As Long
    With myFizz
        GetThree = .ThirdVal
    End With
End Function
";

            var actual = GetReferenceCount(inputCode, "myBizz", "TBizz");
            Assert.AreEqual(2, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void GetsCorrectReferenceCount()
        {
            string inputCode =
$@"
Private Type TBizz
    First As String
    Second As String
End Type

Public Type ToEnsureValidCounts
    First As String
    Second As String
End Type

Private bizz_ As TBizz

Private fizz_ As TBizz

Public Sub Fizz(newValue As String)
    With bizz_
        .First = newValue
    End With
End Sub

Public Sub Buzz(newValue As String)
    With bizz_
        .Second = newValue
    End With
End Sub

Public Sub Fizz1(newValue As String)
    bizz_.First = newValue
End Sub

Public Sub Buzz1(newValue As String)
    bizz_.Second = newValue
End Sub

Public Sub Tazz(newValue As String)
    fizz_.First = newValue
    fizz_.Second = newValue
End Sub
";

            var actual = GetReferenceCount(inputCode, "bizz_", "TBizz");
            Assert.AreEqual(4, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void GetsCorrectReferenceCount_ClassAccessor()
        {
            string className = "TestClass";
            string classCode =
$@"
Public this As TBizz
";

            string classInstance = "theClass";
            string moduleName = MockVbeBuilder.TestModuleName;
            string moduleCode =
$@"
Public Type TBizz
    First As String
    Second As String
End Type

Public Type ToEnsureValidCounts
    First As String
    Second As String
End Type

Private {classInstance} As {className}

Public Sub Initialize()
    Set {classInstance} = New {className}
End Sub

Public Sub Fizz1(newValue As String)
        {classInstance}.this.First = newValue
End Sub

Public Sub Buzz1(newValue As String)
        {classInstance}.this.Second = newValue
End Sub

";

            var actual = GetReferenceCount("this", "TBizz", (moduleName, moduleCode, ComponentType.StandardModule),
                (className, classCode, ComponentType.ClassModule));
            Assert.AreEqual(2, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void SingleElementRefNestedWithStatements()
        {
            string moduleCode =
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

            var actual = GetReferenceCount(moduleCode, "mTypesField", "PType2");
            Assert.AreEqual(2, actual);
        }

        [TestCase("AValue", ".CInst.BInst.AInst.AValue", "mTestField", true)]
        [TestCase("BValue", ".CInst.BInst.BValue", "mTestField", true)]
        [TestCase("CValue", ".CInst.CValue", "mTestField", true)]
        [TestCase("DValue", ".DValue", "mTestField", true)]
        [TestCase("AValue", ".CInst.BInst.AInst.AValue", "mBogeyField", false)]
        [TestCase("BValue", ".CInst.BInst.BValue", "mBogeyField", false)]
        [TestCase("CValue", ".CInst.CValue", "mBogeyField", false)]
        [TestCase("DValue", ".DValue", "mBogeyField", false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void IsRelatedReferenceMemberAccess(string udtMemberID, string memberAccess, string targetID, bool expected)
        {
            string moduleName = MockVbeBuilder.TestModuleName;
            string moduleCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue}

Private mBogeyField As DType

Private mTestField As DType

Sub TestBogey(arg As Long)  'Same udtMember => not the mTestField instance
    mBogeyField{memberAccess} = arg
End Sub

Sub Test(arg As Long)
    mTestField{memberAccess} = arg
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((moduleName, moduleCode, ComponentType.StandardModule));

            var actual = RunIsRelatedReferenceTest(vbe.Object, targetID, udtMemberID);
            Assert.AreEqual(expected, actual);
        }

        [TestCase("AValue", ".CInst.BInst.AInst", "mTestField", true)]
        [TestCase("BValue", ".CInst.BInst", "mTestField", true)]
        [TestCase("CValue", ".CInst", "mTestField", true)]
        [TestCase("DValue", "", "mTestField", true)]
        [TestCase("AValue", ".CInst.BInst.AInst", "mBogeyField", false)]
        [TestCase("BValue", ".CInst.BInst", "mBogeyField", false)]
        [TestCase("CValue", ".CInst", "mBogeyField", false)]
        [TestCase("DValue", "", "mBogeyField", false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void IsRelatedReferenceWithMemberAccess(string udtMemberID, string memberAccess, string targetID, bool expected)
        {
            string moduleName = MockVbeBuilder.TestModuleName;
            string moduleCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue}

Private mBogeyField As DType

Private mTestField As DType

Sub TestBogey(arg As Long)  'Same udtMember => not the mTestField instance
    With mBogeyField{memberAccess}
        .{udtMemberID} = arg
    End With
End Sub

Sub Test(arg As Long)
    With mTestField{memberAccess}
        .{udtMemberID} = arg
    End With
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((moduleName, moduleCode, ComponentType.StandardModule));

            var actual = RunIsRelatedReferenceTest(vbe.Object, targetID, udtMemberID);
            Assert.AreEqual(expected, actual);
        }

        [TestCase("AValue", ".CInst.BInst", "AInst.AValue", "mTestField", true)]
        [TestCase("AValue", ".CInst", "BInst.AInst.AValue", "mTestField", true)]
        [TestCase("AValue", "", "CInst.BInst.AInst.AValue", "mTestField", true)]
        [TestCase("AValue", ".CInst.BInst", "AInst.AValue", "mBogeyField", false)]
        [TestCase("AValue", ".CInst", "BInst.AInst.AValue", "mBogeyField", false)]
        [TestCase("AValue", "", "CInst.BInst.AInst.AValue", "mBogeyField", false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void IsRelatedReferenceWithMemberAccessSplits(string udtMemberID, string withStmtAccess, string memberAccess, string targetID, bool expected)
        {
            string moduleName = MockVbeBuilder.TestModuleName;
            string moduleCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue}

Private mBogeyField As DType

Private mTestField As DType

Sub TestBogey(arg As Long)  'Same udtMember => not the mTestField instance
    With mBogeyField{withStmtAccess}
        .{memberAccess} = arg
    End With
End Sub

Sub Test(arg As Long)
    With mTestField{withStmtAccess}
        .{memberAccess} = arg
    End With
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((moduleName, moduleCode, ComponentType.StandardModule));

            var actual = RunIsRelatedReferenceTest(vbe.Object, targetID, udtMemberID);
            Assert.AreEqual(expected, actual);
        }

        [TestCase("mTestField", true)]
        [TestCase("mBogeyField", false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void IsRelatedReferenceNestedWithMemberAccess(string targetID, bool expected)
        {
            string moduleName = MockVbeBuilder.TestModuleName;
            string moduleCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue}

Private mBogeyField As DType

Private mTestField As DType

Sub TestBogey(arg As Long)  'Same udtMember => not the mTestField instance
    With mBogeyField
        With .CInst.BInst
            .AInst.AValue = arg
        End With
    End With
End Sub

Sub Test(arg As Long)
    With mTestField
        With .CInst.BInst
            .AInst.AValue = arg
        End With
    End With
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((moduleName, moduleCode, ComponentType.StandardModule));

            var actual = RunIsRelatedReferenceTest(vbe.Object, targetID, "AValue");
            Assert.AreEqual(expected, actual);
        }

        [TestCase("mFizz", true)]
        [TestCase("mBodgey", false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void IsRelatedReferenceClassInstanceReferencedInStdModuleNestedWithStmt(string targetID, bool expected)
        {
            var testModuleName = MockVbeBuilder.TestModuleName;

            var testModuleCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue}

Public {targetID} As DType

Public mBogey As DType";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.ClassModule);

            var udtDeclaringModule = "UDTDeclaringModule";
            var udtDeclaringCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue.Replace("Private", "Public")}

";

            var udtDeclaringModuleStdModule = (moduleName: udtDeclaringModule, udtDeclaringCode, ComponentType.StandardModule);

            var referencingModule = "SomeOtherModule";
            var referencingModuleCode =
$@"

Private mInst As {testModuleName}

Sub Initialize()
    Set mInst = new {testModuleName}
End Sub

Sub TestBogey()
    With mInst.mBogey
        With mInst.mFizz.CInst
            .BInst.AInst.AValue = 0
        End With
        .CInst.BInst.AInst.AValue = 3
    End With
End Sub

Sub Test()
    With mInst.mFizz
        With mInst.mBogey.CInst
            .BInst.BValue = 0
        End With
        .CInst.BInst.AInst.AValue = 3
    End With
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule, udtDeclaringModuleStdModule);
            var actual = RunIsRelatedReferenceTest(vbe.Object, targetID, "AValue");
            Assert.AreEqual(expected, actual);
        }

        [TestCase(MockVbeBuilder.TestModuleName)]
        [TestCase("")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void StdModuleValueMemberAccessQualified(string testModuleName)
        {
            var targetID = "mFizz";
            var expected = testModuleName.Length > 0;
            var refQualificationExpression = expected ? $"{testModuleName}." : string.Empty;

            var testModuleCode =
$@"
Option Explicit

Public {targetID} As Integer
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var referencingModule = "SomeOtherModule";
            var referencingModuleCode =
$@"Sub Bazz()
    {refQualificationExpression}{targetID} = 0
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var actual = RunIsModuleQualifiedExternalReferenceTest(vbe.Object, targetID);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void StdModuleValueWithMemberAccessQualified()
        {
            var targetID = "mFizz";
            var testModuleName = MockVbeBuilder.TestModuleName;
            var expected = true;

            var testModuleCode =
$@"
Option Explicit

Public {targetID} As Integer
";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var referencingModule = "SomeOtherModule";
            var referencingModuleCode =
$@"Sub Bazz()
    With {testModuleName}
        .{targetID} = 0
    End With
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var actual = RunIsModuleQualifiedExternalReferenceTest(vbe.Object, targetID);
            Assert.AreEqual(expected, actual);
        }

        [TestCase(MockVbeBuilder.TestModuleName, ".CInst.BInst.AInst.AValue")]
        [TestCase("", ".CInst.BInst.AInst.AValue")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void StdModuleUDTFieldMemberAccessQualified(string testModuleName, string memberQualificationExpression)
        {
            var targetID = "mFizz";
            var expected = testModuleName.Length > 0;
            var refQualificationExpression = expected ? $"{testModuleName}." : string.Empty;

            var testModuleCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue}

Public {targetID} As DType";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var referencingModule = "SomeOtherModule";
            var referencingModuleCode =
$@"Sub Bazz()
    {refQualificationExpression}{targetID}{memberQualificationExpression} = 0
End Sub
";
            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var actual = RunIsModuleQualifiedExternalReferenceTest(vbe.Object, targetID);
            Assert.AreEqual(expected, actual);
        }

        [TestCase(MockVbeBuilder.TestModuleName)]
        [TestCase("")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void StdModuleUDTFieldWithMemberAccessQualified(string testModuleName)
        {
            var targetID = "mFizz";
            var expected = testModuleName.Length > 0;
            var refQualificationExpression = expected ? $"{testModuleName}." : string.Empty;

            var testModuleCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue}

Public {targetID} As DType";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var referencingModule = "SomeOtherModule";
            var referencingModuleCode =
$@"Sub Bazz()
    With {refQualificationExpression}{targetID}
        .CInst.BInst.AInst.AValue = 0
    End With
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var actual = RunIsModuleQualifiedExternalReferenceTest(vbe.Object, targetID);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void StdModuleOtherUDTFieldWithMemberAccess()
        {
            var targetID = "mFizz";
            var testModuleName = MockVbeBuilder.TestModuleName;
            var expected = false;

            var testModuleCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue}

Public {targetID} As DType

Public mBogey As DType";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var referencingModule = "SomeOtherModule";
            var referencingModuleCode =
$@"Sub Bazz()
    With {testModuleName}.mBogey.CInst
        .BInst.AInst.AValue = 0
        {targetID}.CInst.BInst.AInst.AValue = 3 'targetID not module qualified
    End With
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var actual = RunIsModuleQualifiedExternalReferenceTest(vbe.Object, targetID);
            Assert.AreEqual(expected, actual);
        }

        [TestCase(MockVbeBuilder.TestModuleName)]
        [TestCase("")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void StdModuleUDTFieldWithMemberAccessQualifiedMemberAccessExpression(string testModuleName)
        {
            var targetID = "mFizz";
            var expected = testModuleName.Length > 0;
            var refQualificationExpression = expected ? $"{testModuleName}." : string.Empty;

            var testModuleCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue}

Public {targetID} As DType";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.StandardModule);

            var referencingModule = "SomeOtherModule";
            var referencingModuleCode =
$@"Sub Bazz()
    With {refQualificationExpression}{targetID}.CInst
        .BInst.AInst.AValue = 0
    End With
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule);
            var actual = RunIsModuleQualifiedExternalReferenceTest(vbe.Object, targetID);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void ClassInstanceReferencedInStdModule()
        {
            var targetID = "mFizz";
            var testModuleName = MockVbeBuilder.TestModuleName;
            var expected = true;

            var testModuleCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue}

Public {targetID} As DType

Public mBogey As DType";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.ClassModule);

            var udtDeclaringModule = "UDTDeclaringModule";
            var udtDeclaringCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue.Replace("Private", "Public")}
";

            var udtDeclaringModuleStdModule = (moduleName: udtDeclaringModule, udtDeclaringCode, ComponentType.StandardModule);

            var referencingModule = "SomeOtherModule";
            var referencingModuleCode =
$@"

Private mInst As {testModuleName}

Sub Initialize()
    Set mInst = new {testModuleName}
End Sub

Sub Bazz()
    With mInst.mBogey.CInst
        .BInst.AInst.AValue = 0
        mInst.mFizz.CInst.BInst.AInst.AValue = 3
    End With
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule, udtDeclaringModuleStdModule);
            var actual = RunIsModuleQualifiedExternalReferenceTest(vbe.Object, targetID);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUtilities))]
        public void ClassInstanceReferencedInStdModuleNestedWith()
        {
            var targetID = "mFizz";
            var testModuleName = MockVbeBuilder.TestModuleName;
            var expected = true;

            var testModuleCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue}

Public {targetID} As DType

Public mBogey As DType";

            var declaringModule = (testModuleName, testModuleCode, ComponentType.ClassModule);

            var udtDeclaringModule = "UDTDeclaringModule";
            var udtDeclaringCode =
$@"
Option Explicit

{privateNestedTypeDeclaration_DTypeContainsCTypeContainsBTypeContainsATypeContainsAValue.Replace("Private", "Public")}
";

            var udtDeclaringModuleStdModule = (moduleName: udtDeclaringModule, udtDeclaringCode, ComponentType.StandardModule);

            var referencingModule = "SomeOtherModule";
            var referencingModuleCode =
$@"

Private mInst As {testModuleName}

Sub Initialize()
    Set mInst = new {testModuleName}
End Sub

Sub Bazz()
    With mInst.mFizz
        With mInst.mBogey.CInst
            .BInst.AInst.AValue = 0
            .CInst.BInst.AInst.AValue = 3
        End With
    End With
End Sub
";

            var referencingModuleStdModule = (moduleName: referencingModule, referencingModuleCode, ComponentType.StandardModule);

            var vbe = MockVbeBuilder.BuildFromModules(declaringModule, referencingModuleStdModule, udtDeclaringModuleStdModule);
            var actual = RunIsModuleQualifiedExternalReferenceTest(vbe.Object, targetID);
            Assert.AreEqual(expected, actual);
        }

        private static bool RunIsModuleQualifiedExternalReferenceTest(IVBE vbe, string targetID)
        {
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var this_Target = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Where(d => d.IdentifierName == targetID)
                    .Single();

                return EncapsulateFieldUtilities.IsModuleQualifiedExternalReferenceOfUDTField(state, this_Target.References.FirstOrDefault(), this_Target.QualifiedModuleName);
            }
        }

        private static bool RunIsRelatedReferenceTest(IVBE vbe, string targetID, string udtMemberID, string precedingSubID = "Test")
        {
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var this_Target = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Where(d => d.IdentifierName == targetID).OfType<VariableDeclaration>()
                    .Single();

                var testSub = state.DeclarationFinder.UserDeclarations(DeclarationType.Procedure)
                    .Where(d => d.IdentifierName == precedingSubID)
                    .Single();

                var targetRef = state.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                    .Where(udtM => udtM.IdentifierName == udtMemberID)
                    .SelectMany(d => d.References)
                    .Where(rf => rf.Selection > testSub.Selection) //the 'targetID' answer is the reference after the preceding SubID
                    .Select(rf => rf).FirstOrDefault();

                return EncapsulateFieldUtilities.IsRelatedUDTMemberReference(this_Target, targetRef);
            }
        }
        private static long GetReferenceCount(string inputCode, string targetID, string udtID)
        {
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, inputCode));
            return GetReferenceCount(vbe.Object, targetID, udtID);
        }

        private static long GetReferenceCount(string targetID, string udtID, params (string moduleName, string content, ComponentType compType)[] modules)
        {
            var vbe = MockVbeBuilder.BuildFromModules(modules);
            return GetReferenceCount(vbe.Object, targetID, udtID);
        }
        private static long GetReferenceCount(IVBE vbe, string targetID, string udtID)
        {
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var target = GetFieldDeclaration(state, targetID);
                var allUDTMemberRefs = GetUDTMemberReferences(state, udtID).SelectMany(d => d.References);

                var udtMemberRefs = allUDTMemberRefs.Where(rf => EncapsulateFieldUtilities.IsRelatedUDTMemberReference(target, rf))
                    .Select(rf => rf);

                return udtMemberRefs.Count();
            }
        }

        private static VariableDeclaration GetFieldDeclaration(IDeclarationFinderProvider declarationFinderProvider, string identifier)
        {
            return declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .OfType<VariableDeclaration>()
                .Where(d => identifier == d.IdentifierName)
                .Single();
        }

        private static IEnumerable<Declaration> GetUDTMemberReferences(IDeclarationFinderProvider declarationFinderProvider, string identifier)
        {
            return declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                .Where(d => identifier == d.ParentDeclaration.IdentifierName);
        }
    }
}
