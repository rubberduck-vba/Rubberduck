using NUnit.Framework;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using TestResolver = RubberduckTests.Refactoring.ConflictDetectorTestsResolver;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ConflictDetectionSessionRenameTests
    {
        //MS_VBAL 5.4.3.1, 5.4.3.2
        //Method names need to be different than contained local variables or constants
        [TestCase("Dim goo As Long", "GOO", true)]
        [TestCase("Static goo As Long", "GOO", true)]
        [TestCase("Const goo As Long = 10", "GOO", true)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void ProcedureNameChangeConflictsWithLocalDeclaration(string declaration, string newName, bool expected)
        {
            var selection = ("Fizz", DeclarationType.Procedure);
            var inputCode =
$@"
Option Explicit

Public Sub Fizz(arg As Long)
    {declaration}
End Sub
";
            Assert.AreEqual(expected, TestForRenameConflict(inputCode, selection, newName));
        }

        //MS_VBAL 5.4.3.1, 5.4.3.2
        //Method names need to be different than containedlocal variables or constants
        [TestCase("Dim goo As Long", "GOO", true)]
        [TestCase("Static goo As Long", "GOO", true)]
        [TestCase("Const goo As Long = 10", "GOO", true)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void FunctionNameChangeConflictsWithLocalDeclaration(string declaration, string newName, bool expected)
        {
            var selection = ("Fizz", DeclarationType.Function);
            var inputCode =
$@"
Option Explicit

Public Function Fizz(arg As Long) As Long
    {declaration}
End Function
";
            Assert.AreEqual(expected, TestForRenameConflict(inputCode, selection, newName));
        }

        //MS_VBAL 5.4.3.1, 5.4.3.2
        //Method names need to be different than containedlocal variables or constants
        [TestCase("Dim goo As Long", "GOO", true)]
        [TestCase("Static goo As Long", "GOO", true)]
        [TestCase("Const goo As Long = 10", "GOO", true)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void PropertyNameChangeConflictsWithLocalDeclaration(string declaration, string newName, bool expected)
        {
            var selection = ("Fizz", DeclarationType.PropertyGet);
            var inputCode =
$@"
Option Explicit

Private mFizz As Long

Public Property Get Fizz() As Long
    {declaration}
    Fizz = mFizz 
End Property
";
            Assert.AreEqual(expected, TestForRenameConflict(inputCode, selection, newName));
        }

        //MS-VBAL 5.3.1.5 - parameters cannot have the same name as their containing Function
        //So, test for renaming a Function to match an existing parameter
        [TestCase("Fizz", DeclarationType.PropertyGet, "arg", true)]
        [TestCase("Fizz", DeclarationType.PropertyLet, "arg", false)]
        [TestCase("Fizz", DeclarationType.PropertyLet, "value", false)]
        [TestCase("Gazz", DeclarationType.Function, "gazzArg", true)]
        [TestCase("Guzz", DeclarationType.Procedure, "guzzArg", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void FunctionRenameConflictsWithParameters(string target, DeclarationType declarationType, string newName, bool expected)
        {
            var selection = (target, declarationType);
            var inputCode =
$@"
Option Explicit

Private mFizz As Long
Private mGuzz As String

Public Property Get Fizz(arg As Long) As Long
    Fizz = mFizz /arg 
End Property

Public Property Let Fizz(arg As Long, value As Long)
    mFizz = value * arg
End Property

Public Function Gazz(gazzArg As String) As String
    Gazz = ""asdf"" 'gazzArg + ""2""
End Function

Public Sub Guzz(guzzArg As String)
    mGuzz = guzzArg + ""2""
End Sub
";
            Assert.AreEqual(expected, TestForRenameConflict(inputCode, selection, newName));
        }

        //Edge-case: MS-VBAL 5.3.1.5 (parameters cannot have the same name as their containing Function)
        //would imply that renaming a Subroutine to one of its parameter's identifiers should be OK. 
        //And, it is, for many scenarios.
        //But, in a self-referential/recursive scenario, uncompilable code can occur.
        [Test]
        [Category(nameof(ConflictDetectionSession))]
        public void SubRenameConflictsWithParameterRecursive()
        {
            var selection = ("Guzz", DeclarationType.Procedure);
            var inputCode =
$@"
Option Explicit

Private mGuzz As String
Private mCount As Long

Public Sub Guzz(guzzArg As String)
    mCount = mCount + 1
    mGuzz = guzzArg + CStr(mCount)
    If mCount < 6 Then Guzz mGuzz
End Sub
";
            Assert.AreEqual(true, TestForRenameConflict(inputCode, selection, "guzzArg"));
        }

        //MS_VBAL 5.3.1.6: each subroutine and Function name must be different than
        //any other module variable, module Constant, EnumerationMember, or Procedure
        //defined in the same module
        [TestCase("mFazz", true)]
        [TestCase("constFazz", true)]
        [TestCase("Bazz", true)]
        [TestCase("Fazz", true)]
        [TestCase("SecondValue", true)]
        [TestCase("Gazz", false)]
        [TestCase("ETest", false)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void MethodRenameChangeConflicts(string newName, bool expected)
        {
            var selection = ("Fizz", DeclarationType.Function);
            var inputCode =
$@"
Option Explicit

Public Enum ETest
    FirstValue = 0
    SecondValue
End Enum

Private mFazz As String

Private Const constFazz As Long = 7

Private mFizz As Long

Public Function Fizz() As Long
    Fizz = mFizz
End Function

Public Sub Bazz() 
End Sub

Public Property Get Fazz() As Long
    Fazz = mFazz
End Property

Public Property Let Fazz(value As Long)
    mFazz =  value
End Property
";
            Assert.AreEqual(expected, TestForRenameConflict(inputCode, selection, newName));
        }

        [TestCase("mFazz", true)]
        [TestCase("constFazz", true)]
        [TestCase("Fizz", true)]
        [TestCase("Bazz", true)]
        [TestCase("Fazz", true)]
        [TestCase("SecondValue", true)]
        [TestCase("ETest", false)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void EnumMemberRenameChangeConflicts(string newName, bool expected)
        {
            var selection = ("FirstValue", DeclarationType.EnumerationMember);
            var inputCode =
$@"
Option Explicit

Public Enum ETest
    FirstValue = 0
    SecondValue
End Enum

Private mFazz As String

Private Const constFazz As Long = 7

Private mFizz As Long

Public Function Fizz() As Long
    Fizz = mFizz
End Function

Public Sub Bazz() 
End Sub

Public Property Get Fazz() As Long
    Fazz = mFazz
End Property

Public Property Let Fazz(value As Long)
    mFazz =  value
End Property
";
            Assert.AreEqual(expected, TestForRenameConflict(inputCode, selection, newName));
        }

        //MS_VBAL 5.3.1.7: 
        //Each property Let\Set\Get must have a unique name
        [TestCase("Fazz", DeclarationType.PropertyGet, "Fizz", false)]
        [TestCase("Fazz", DeclarationType.PropertyLet, "Fizz", true)]
        [TestCase("Fazz", DeclarationType.PropertyGet, "Fuzz", true)]
        [TestCase("Fazz", DeclarationType.PropertySet, "Fuzz", false)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void RenameLetSetGetAreUnique(string targetName, DeclarationType targetDeclarationType, string newName, bool expected)
        {
            var selection = (targetName, targetDeclarationType);
            var inputCode =
$@"
Option Explicit


Private mFazz As Variant
Private mFizz As Variant
Private mFuzz As Variant

Public Property Get Fazz() As Variant
    If IsObject(mFazz) Then
        Set Fazz = mFazz
    Else
        Fazz = mFazz
    Endif
End Property

Public Property Let Fazz(value As Variant)
    mFazz =  value
End Property

Public Property Set Fazz(value As Variant)
    Set mFazz =  value
End Property

Public Property Let Fizz(value As Variant)
    mFizz =  value
End Property

Public Property Get Fuzz() As Variant
    If IsObject(mFuzz) Then
        Set Fuzz = mFuzz
    Else
        Fuzz = mFuzz
    Endif
End Property
";
            Assert.AreEqual(expected, TestForRenameConflict(inputCode, selection, newName));
        }

        //MS_VBAL 5.3.1.7: 
        //Each shared name must have the same asType, parameters, etc
        [TestCase("(value As Long)", "()", false)]
        [TestCase("(value As Variant)", "()", true)]  //Inconsistent AsTypeName
        [TestCase("(value As Long)", "(arg1 As String)", true)] //Inconsistent parameters (quantity)
        [TestCase("(arg1 As Boolean, value As Long)", "(arg1 As String)", true)] //Inconsistent parameters (type)
        [TestCase("(ByVal arg1 As String, value As Long)", "(arg1 As String)", true)] //Inconsistent parameters (parameter mechanism)
        [TestCase("(arg1 As String, arg2 As Long, value As Long)", "(arg1 As String, arg22 As Long)", true)] //Inconsistent parameters (parameter name)
        [TestCase("(arg1 As String, arg2 As Variant, value As Long)", "(arg1 As String, ParamArray arg2() As Variant)", true)] //Inconsistent parameters (ParamArray)
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void RenamePropertyInconsistentSignatures(string letParameters, string getParameters, bool expected)
        {
            var sourceContent =
$@"
Option Explicit

Private mFizz As Long
Private mFazz As Long

Public Property Let Fi|zz{letParameters}
    mFizz =  value
End Property

Public Property Get Fazz{getParameters} As Long
    mFazz =  value
End Property
";
            Assert.AreEqual(expected, TestForRenameConflict("Fazz", sourceContent));
        }

        [TestCase("multiplier", true)]
        [TestCase("adder", true)]
        [TestCase("Foo", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void NewFunctionNameConflictsWithLocalVariable(string newName, bool isConflict)
        {
            var moduleContent1 =
            @"

Private member1 As Long

Public Function F|oo(arg As Long) As Long
    Dim adder as Long
    Const multiplier As Long = 10
    adder = 10
    Foo = (arg + adder) * multiplier
End Function
";
            var hasConflict = TestForRenameConflict(newName, moduleContent1);
            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("Fizz", true)]
        [TestCase("Foo", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void NewFunctionNameConflictsWithReferenceInOtherMember(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
            @"

Private member1 As Long

Public Function F|oo(arg As Long) As Long
    Foo = arg * 2
End Function

Public Sub Goo(arg As Long)
    Dim fizz As Long
    member1 = Foo(arg + fizz)
End Sub
";
            var hasConflict = TestForRenameConflict(nameToCheck, moduleContent1);
            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("Fizz() As Long", DeclarationType.Function, ComponentType.StandardModule, true)]
        [TestCase("Fizz()", DeclarationType.Procedure, ComponentType.StandardModule, true)]
        [TestCase("Let Fizz(value As Long)", DeclarationType.PropertyLet, ComponentType.StandardModule, true)]
        [TestCase("Set Fizz(value As Variant)", DeclarationType.PropertySet, ComponentType.StandardModule, true)]
        [TestCase("Get Fizz() As Long", DeclarationType.PropertyGet, ComponentType.StandardModule, true)]
        [TestCase("Fizz() As Long", DeclarationType.Function, ComponentType.ClassModule, false)]
        [TestCase("Fizz()", DeclarationType.Procedure, ComponentType.ClassModule, false)]
        [TestCase("Let Fizz(value As Long)", DeclarationType.PropertyLet, ComponentType.ClassModule, false)]
        [TestCase("Set Fizz(value As Variant)", DeclarationType.PropertySet, ComponentType.ClassModule, false)]
        [TestCase("Get Fizz() As Long", DeclarationType.PropertyGet, ComponentType.ClassModule, false)]
        [TestCase("Fizz()", DeclarationType.Procedure, ComponentType.UserForm, false)]
        [Category(nameof(ConflictDetectionSession))]
        public void RenamedMemberConflictsWithExternallyReferencedPublicModuleScopeEntities(string declaration, DeclarationType declarationType, ComponentType testModuleComponentType, bool expected)
        {
            var memberType = Tokens.Property;
            var signature = string.Empty;
            switch (declarationType)
            {
                case DeclarationType.Function:
                    memberType = Tokens.Function;
                    signature = $"{memberType} {declaration}";
                    break;
                case DeclarationType.Procedure:
                    memberType = Tokens.Sub;
                    signature = $"{memberType} {declaration}";
                    break;
                case DeclarationType.PropertyLet:
                    signature = $"{memberType} {declaration}";
                    break;
                case DeclarationType.PropertySet:
                    signature = $"{memberType} {declaration}";
                    break;
                case DeclarationType.PropertyGet:
                    signature = $"{memberType} {declaration}";
                    break;
            }

            var conflictName = "testEntity";
            var testModuleCode =
$@"
Option Explicit

Public {signature}
End {memberType}

";
            var referencingModuleCode =
$@"
Option Explicit

Public Function DoIt() As Long
    DoIt = {conflictName}
End Function
";

            var publicFieldModuleCode =
$@"
Option Explicit

Public {conflictName} As Long

";

            var vbe = MockVbeBuilder.BuildFromModules(
                (MockVbeBuilder.TestModuleName, testModuleCode, testModuleComponentType),
                ("ReferencingModule", referencingModuleCode, ComponentType.StandardModule),
                ("PublicFieldModule", publicFieldModuleCode, ComponentType.StandardModule));
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var target = state.DeclarationFinder.MatchName("Fizz").Where(d => d.DeclarationType.HasFlag(declarationType)).Single();

                var hasConflict = HasRenameConflict(state, target, conflictName);
                Assert.AreEqual(expected, hasConflict);
            }
        }

        [TestCase("member1", true)]
        [TestCase("adder", false)]
        [TestCase("Foo", true)]
        [Category(nameof(ConflictDetectionSession))]
        public void RenameLocalVariableInFunction(string newName, bool isConflict)
        {
            var moduleContent1 =
            @"

Private member1 As Long

Public Function Foo() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + ad|der
    Foo = member1
End Function
";

            var hasConflict = TestForRenameConflict(newName, moduleContent1);
            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("member1", true)]
        [TestCase("adder", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void RenameFunctionToLocalVariable(string newName, bool isConflict)
        {
            var moduleContent1 =
            @"

Private member1 As Long

Public Function Fi|zz() As Long
    Fizz = member1
End Function

Public Function Goo() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder
    Goo = member1
End Function
";

            var hasConflict = TestForRenameConflict(newName, moduleContent1);
            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("member1", false)]
        [TestCase("adder", true)]
        [TestCase("Foo", true)]
        [Category(nameof(ConflictDetectionSession))]
        public void InProcedure_MemberDeclaration(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
@"
Private member1 As Long

Public Function Foo() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder
    Foo = m|ember1
End Function
";
            var hasConflict = TestForRenameConflict(nameToCheck, moduleContent1);
            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("member1", false)]
        [TestCase("member2", false)]
        [TestCase("adder", true)]
        [TestCase("Foo", true)]
        [TestCase("Foo2", false)]
        [TestCase("Bar", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void ModuleScope(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
            @"

Private member1 As Long

Public Function Foo() As Long
    Dim adder as Long
    adder = 10
    member1 = memb|er1 + adder
    Foo = member1
End Function
";
            var moduleContent2 =
            @"

Private member1 As Long
Public member2 As Long

Public Function Foo2() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder
    Foo2 = member1
End Function

Private Sub Bar()
    member2 = member2 * 4
End Sub
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("CFirstClass", selection, nameToCheck,
                    ("CFirstClass", ComponentType.ClassModule, code),
                    ("modOne", ComponentType.StandardModule, moduleContent2));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("Foo", false)]
        [TestCase("Foo2", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void PublicClassAndPublicModuleSub_RenameClassSub(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
            @"
Public Function Fo|o() As Long
    Foo = 5
End Function
";
            var moduleContent2 =
            @"
Public Function Foo2() As Long
    Foo2 = 2
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("CFirstClass", selection, nameToCheck,
                    ("CFirstClass", ComponentType.ClassModule, code),
                    ("modOne", ComponentType.StandardModule, moduleContent2));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("Bazz", true)]
        //[TestCase("Gazz", true)]
        //[TestCase("Buzz", true)]
        //[TestCase("Fooz", true)]
        //[TestCase("Tizz", true)]
        [TestCase("Bazz2", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void RenamedFunctionReferenceConflictsWithOtherModuleMember(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
            @"
Public Function Fi|zz() As Long
    Fizz = 10
End Function
";
            var moduleContent2 =
@"

Public Fooz As Long
Private Gazz As Long

Private Const Buzz As Long = 10

Public Function Bazz() As Long
    Dim tizz As Long
    tizz = 25
    Gazz = Fizz + Buzz + tizz
    Bazz = Gazz
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modTwo", ComponentType.StandardModule, moduleContent2));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("Gazz", true)]
        [TestCase("Buzz", true)]
        [TestCase("Fooz", true)]
        [TestCase("Bazz2", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void RenamedFunctionReferenceConflictsWithModuleScopeEntities(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
            @"
Public Function Fi|zz() As Long
    Fizz = 10
End Function
";
            var moduleContent2 =
@"

Public Fooz As Long
Private Gazz As Long

Private Const Buzz As Long = 10

Public Function Bazz() As Long
    Dim tizz As Long
    tizz = 25
    Gazz = Fizz + Buzz + tizz
    Bazz = Gazz
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modTwo", ComponentType.StandardModule, moduleContent2));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("Tizz", true)]
        [TestCase("Bazz2", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void RenamedFunctionReferenceConflictsWithLocalVariable(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
            @"
Public Function Fi|zz() As Long
    Fizz = 10
End Function
";
            var moduleContent2 =
@"

Public Fooz As Long
Private Gazz As Long

Private Const Buzz As Long = 10

Public Function Bazz() As Long
    Dim tizz As Long
    tizz = 25
    Gazz = Fizz + Buzz + tizz
    Bazz = Gazz
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modTwo", ComponentType.StandardModule, moduleContent2));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("Bazz", true)]
        [TestCase("Buzz", true)]
        [TestCase("Fooz", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void RenamedFunctionConflicts(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
            @"
Public Function Fi|zz() As Long
    Fizz = Bazz(Buzz)
End Function
";
            var moduleContent2 =
@"

Public Fooz As Long
Private Gazz As Long

Public Const Buzz As Long = 10

Public Function Bazz(arg As Long) As Long
    Dim tizz As Long
    tizz = arg
    Gazz = Buzz + tizz
    Bazz = Gazz
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modTwo", ComponentType.StandardModule, moduleContent2));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [Test]
        [TestCase("Foo", false)]
        [TestCase("Foo2", true)]
        [TestCase("member11", true)]
        [TestCase("member1", false)]
        [TestCase("Bar", true)]
        [TestCase("adder", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void Module_To_ClassScope(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
            @"

Private member11 As Long
Public member2 As Long

Public Function Foo2() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder
    Foo2 = member1
End Function

Private Sub Bar()
    member2 = membe|r2 * 4
End Sub
";
            var moduleContent2 =
            @"

Private member1 As Long

Public Function Foo() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder
    Foo = member1
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("CFirstClass", selection, nameToCheck,
                    ("CFirstClass", ComponentType.ClassModule, code),
                    ("modOne", ComponentType.StandardModule, moduleContent2));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("theFilename", true)] //Conflicts with parameter identifier
        [TestCase("member1", true)] //Conflicts with module variable identifier
        [TestCase("TestSub", true)] //Conflicts with module procedure
        [TestCase("TestFunction", true)] //Conflicts with module function
        [Category(nameof(ConflictDetectionSession))]
        public void PrivateSub_SameModuleConflicts(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
@"
Private member1 As String

Public Function GetFil|ename(theFilename As String) As String
    GetFilename = member1
End Function

Private Sub TestSub()
End Sub

Private Function TestFunction()
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code));

            Assert.AreEqual(isConflict, hasConflict);
        }


        [TestCase("theFilename", false)] //No conflicts with parameter identifier unless it is a function
        [Category(nameof(ConflictDetectionSession))]
        public void ProcedureNameMatchesParameter_NoConflict(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
@"
Private member1 As String

Public Sub SetFil|ename(theFilename As String)
    member1 = theFilename
End Sub
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("ExtractFilename", true)] //NonQualified reference conflicts with function
        [TestCase("StoreFilename", true)] //NonQualified reference conflicts with procedure
        [Category(nameof(ConflictDetectionSession))]
        public void PublicSub_NonQualifiedReferenceConflictsWithMembersInOtherModules(string nameToCheck, bool isConflict)
        {

            var moduleContent1 =
@"
Private member1 As String

Public Sub SetFil|ename(filename As String)
    member1 = filename
End Sub
";
            var moduleContent3 =
@"
Private mFolderpath As String

Private Sub StoreFilename(filepath As String)
    Dim theFileName As String
    theFileName = ExtractFilename(filepath)
    SetFilename theFileName
End Sub

Private Function ExtractFilename(filepath As String) As String
    Dim cache As String
    cache = filepath
    ExtractFilename = filepath
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modThree", ComponentType.StandardModule, moduleContent3));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("mFolderpath", true)] //NonQualified reference conflicts with Field identifier
        [Category(nameof(ConflictDetectionSession))]
        public void PublicSub_NonQualifiedReferenceConflictsWithModuleVariablesInOtherModules(string nameToCheck, bool isConflict)
        {

            var moduleContent1 =
@"
Public Sub SetFil|ename(value As String)
End Sub
";
            var moduleContent3 =
@"
Private mFolderpath As String

Private Sub StoreFilename(filepath As String)
    Dim theFileName As String
    theFileName = ExtractFilename(filepath)
    SetFilename theFileName
End Sub

Private Function ExtractFilename(filepath As String) As String
    Dim cache As String
    cache = filepath
    ExtractFilename = filepath
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modThree", ComponentType.StandardModule, moduleContent3));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("theFileName", true)] //Reference conflicts with localVariable
        [TestCase("cache", false)] //No conflicts with matching local variable in different scope
        [Category(nameof(ConflictDetectionSession))]
        public void PublicSub_NonQualifiedReferenceConflictsWithLocalVariablesInOtherModules(string nameToCheck, bool isConflict)
        {

            var moduleContent1 =
@"
Public Sub SetFil|ename(value As String)
End Sub
";
            var moduleContent3 =
@"
Private mFolderpath As String

Private Sub StoreFilename(filepath As String)
    Dim theFileName As String
    theFileName = ExtractFilename(filepath)
    SetFilename theFileName
End Sub

Private Function ExtractFilename(filepath As String) As String
    Dim cache As String
    cache = filepath
    ExtractFilename = filepath
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modThree", ComponentType.StandardModule, moduleContent3));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("theFileExtension", true)] //Reference conflicts with localConstant
        [TestCase("cache", false)] //No conflicts with matching local variable in different scope
        [TestCase("mFolderpath", true)] //NonQualified reference conflicts with Module Constant identifier
        [Category(nameof(ConflictDetectionSession))]
        public void PublicSub_NonQualifiedReferenceConflictsWithConstantsInOtherModules(string nameToCheck, bool isConflict)
        {

            var moduleContent1 =
@"
Public Sub SetFil|ename(value As String)
End Sub
";
            var moduleContent3 =
@"
Private Const mFolderpath As String = ""C:\Test\""

Private Sub StoreFilename(filepath As String)
    Const theFileExtension As String = "".xlsb""
    theFileName = ExtractFilename(filepath)
    SetFilename theFileName
End Sub

Private Function ExtractFilename(filepath As String) As String
    Const cache As String = ""Not Used""
    ExtractFilename = filepath
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modThree", ComponentType.StandardModule, moduleContent3));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("filepath", true)] //NonQualified reference conflicts with parameter
        [Category(nameof(ConflictDetectionSession))]
        public void PublicSub_NonQualifiedReferenceConflictsWithParameterInOtherModules(string nameToCheck, bool isConflict)
        {

            var moduleContent1 =
@"
Public Sub SetFil|ename(value As String)
End Sub
";
            var moduleContent3 =
@"
Private mFolderpath As String

Private Sub StoreFilename(filepath As String)
    Dim theFileName As String
    theFileName = ExtractFilename(filepath)
    SetFilename theFileName
End Sub

Private Function ExtractFilename(otherFilepath As String) As String
    ExtractFilename = otherFilepath
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modThree", ComponentType.StandardModule, moduleContent3));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("Bar", true)]
        [TestCase("myData", false)]
        [TestCase("mDupData", false)]
        [TestCase("filepath", false)]
        [TestCase("helper", false)]
        [TestCase("CFileHelper", false)]
        [TestCase("filename", true)]
        [TestCase("mFolderpath", true)]
        [TestCase("ExtractFilename", true)]
        [TestCase("SetFilename", true)]
        [TestCase("Foo", false)]
        [TestCase("FooBar", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void PrivateSub_WithAccessBlock(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
@"
Private myData As String
Private mDupData As String

Public Sub Foo(filenm As String)
    Dim filepath As String
    filepath = ""C:\MyStuff\"" & filenm
    Dim helper As CFileHelper
    Set helper = new CFileHelper
    With helper
        .StoreFile|name filepath
        mDupData = filepath
    End With
End Sub

Public Sub StoreFilename(filename As String)
    myData = filename
End Sub

Private Sub FooBar()
End Sub
";
            var moduleContent2 =
@"
Private mFolderpath As String
Private mFilepath As String

Public Sub StoreFilename(arg As String)
    Dim filename As String
    filename = ExtractFilename(arg)
    SetFilename filename
End Sub

Private Function ExtractFilename(filepath As String) As String
    ExtractFilename = filepath
End Function

Private Sub SetFilename(arg As String)
    mFilepath = arg
End Sub

Public Sub Bar()
End Sub
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("CFileHelper", ComponentType.ClassModule, moduleContent2));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("Foo1", true, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("Foo2", false, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("Foo3", true, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("Foo4", true, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("gConstant", true, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("member2", false, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("member11", false, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("gConstant", true, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("Bar1", false, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("Bar1", true, "Foo1 + Fo|o2 + Foo3")]
        [TestCase("Bar2", false, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("Bar2", true, "Foo1 + Fo|o2 + Foo3")]
        [Category(nameof(ConflictDetectionSession))]
        public void Module_To_ModuleScopeResolutions(string nameToCheck, bool isConflict, string scopeResolvedInput)
        {
            var moduleContent1 =
$@"
Private member11 As Long
Public member2 As Long

Private Function Bar1() As Long
    Bar2
    Bar1 = member2 + {scopeResolvedInput}
End Function

Private Sub Bar2()
    member2 = member2 * 4 
End Sub
";
            var moduleContent2 =
@"
Public Const gConstant As Long = 10

Public Function Foo1() As Long
    Foo1 = 1
End Function

Public Function Foo2() As Long
    Foo2 = 2
End Function

Public Function Foo3() As Long
    Foo3 = 3
End Function

Private Sub Foo4()
End Sub
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modTwo", ComponentType.StandardModule, moduleContent2));

            Assert.AreEqual(isConflict, hasConflict);
        }


        //https://github.com/rubberduck-vba/Rubberduck/issues/4969
        [Test]
        [Category(nameof(ConflictDetectionSession))]
        public void NameConflictDetectionRespectsProjectScope()
        {
            var projectTwoModule = "ProjectTwoModule"; //try to rename ProjectOneModule to this
            var renameTargetModule = "ProjectOneModule";

            var moduleContent = $"Private Sub Foo(){Environment.NewLine}End Sub";

            var projects = new Dictionary<string, IEnumerable<(string, string, ComponentType)>>()
            {
                ["ProjectOne"] = new List<(string, string, ComponentType)>() { (renameTargetModule, moduleContent, ComponentType.StandardModule) },
                ["ProjectTwo"] = new List<(string, string, ComponentType)>() { (projectTwoModule, moduleContent, ComponentType.StandardModule), }
            };

            var builder = new MockVbeBuilder();
            foreach (var project in projects.Keys)
            {
                builder = AddProject(builder, project, projects[project]);
            }
            var vbe = builder.Build().Object;

            using (var parser = MockParser.CreateAndParse(vbe))
            {
                var target = parser.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule)
                    .FirstOrDefault(item => item.IdentifierName.Equals(renameTargetModule));

                var hasConflict = HasRenameConflict(parser, target, projectTwoModule);
                Assert.IsFalse(hasConflict);
            }
        }

        [TestCase("Type", "Blah As Long", DeclarationType.UserDefinedType)]
        [TestCase("Enum", "Blah", DeclarationType.Enumeration)]
        [Category(nameof(ConflictDetectionSession))]
        public void UDTRenameRespectsOtherProjectName(string udtOrEnum, string members, DeclarationType declarationType)
        {
            var projectTwoIdentifier = "ProjectTwo";
            var moduleContent =
$@"
Option Explicit

Public {udtOrEnum} SUT
    {members}
End {udtOrEnum}
";

            var projects = new Dictionary<string, IEnumerable<(string, string, ComponentType)>>()
            {
                [MockVbeBuilder.TestProjectName] = new List<(string, string, ComponentType)>() { (MockVbeBuilder.TestModuleName, moduleContent, ComponentType.StandardModule) },
                [projectTwoIdentifier] = new List<(string, string, ComponentType)>() { (MockVbeBuilder.TestModuleName, "Option Explicit", ComponentType.StandardModule), }
            };

            var builder = new MockVbeBuilder();
            foreach (var project in projects.Keys)
            {
                builder = AddProject(builder, project, projects[project]);
            }
            var vbe = builder.Build().Object;

            using (var parser = MockParser.CreateAndParse(vbe))
            {
                var target = parser.DeclarationFinder.DeclarationsWithType(declarationType).Single();

                var hasConflict = HasRenameConflict(parser, target, projectTwoIdentifier);
                Assert.IsTrue(hasConflict);
            }
        }


        [TestCase("Type", "Blah As Long")]
        [TestCase("Enum", "Blah")]
        [Category(nameof(ConflictDetectionSession))]
        public void ProjectRenameRespectsUDTEnumInOtherProject(string udtOrEnum, string members)
        {
            var projectToRename = "ProjectTwo";
            var testConflictIdentifier = "BlahBlah";
            var moduleContent =
$@"
Option Explicit

Public {udtOrEnum} {testConflictIdentifier}
    {members}
End {udtOrEnum}
";

            var projects = new Dictionary<string, IEnumerable<(string, string, ComponentType)>>()
            {
                [MockVbeBuilder.TestProjectName] = new List<(string, string, ComponentType)>() { (MockVbeBuilder.TestModuleName, moduleContent, ComponentType.StandardModule) },
                [projectToRename] = new List<(string, string, ComponentType)>() { (MockVbeBuilder.TestModuleName, "Option Explicit", ComponentType.StandardModule), }
            };

            var builder = new MockVbeBuilder();
            foreach (var project in projects.Keys)
            {
                builder = AddProject(builder, project, projects[project]);
            }
            var vbe = builder.Build().Object;

            using (var parser = MockParser.CreateAndParse(vbe))
            {
                var target = parser.DeclarationFinder.DeclarationsWithType(DeclarationType.Project)
                    .Single(p => p.IdentifierName.Equals(projectToRename));

                var hasConflict = HasRenameConflict(parser, target, testConflictIdentifier);
                Assert.IsTrue(hasConflict);
            }
        }

        [Test]
        [Category(nameof(ConflictDetectionSession))]
        public void RenameModuleToExistingModuleIdentifierConflict()
        {
            var nameConflictModule = "modTwo";


            (string code, Selection selection) = ToCodeAndSelectionTuple("Option Explicit");

            var hasConflict = TestForRenameConflict("modOne", selection, nameConflictModule,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modTwo", ComponentType.StandardModule, "Option Explicit"));

            Assert.IsTrue(hasConflict);
        }

        [Test]
        [Category(nameof(ConflictDetectionSession))]
        public void ModuleNameMatchesProjectName()
        {
            (string code, Selection selection) = ToCodeAndSelectionTuple("Option Explicit");

            var hasConflict = TestForRenameConflict("modOne", selection, MockVbeBuilder.TestProjectName,
                    ("modOne", ComponentType.StandardModule, code));

            Assert.IsTrue(hasConflict);
        }

        [TestCase("Bazz", true)]
        [TestCase("Tazz", false)]
        [TestCase("mTazz", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void RenameEvent(string newName, bool isConflict)
        {
            var moduleContent1 =
$@"
Public Event Fi|zz(ByVal arg1 As Integer, ByVal arg2 As String)

Public Event Bazz(ByVal arg1 As String, ByVal arg2 As String)

Private mTazz As Long

Public Sub Tazz(arg As Long)
    mTazz = arg
End Sub
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("ClassOne", selection, newName,
                    ("ClassOne", ComponentType.ClassModule, code));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("arg2", true)]
        [TestCase("Tazz", false)]
        [TestCase("mTazz", true)] //Does not violate MS-VBAL conflict rules, but will change logic 
        [Category(nameof(ConflictDetectionSession))]
        public void RenameSubroutineParameter(string newName, bool isConflict)
        {
            var moduleContent1 =
$@"

Private mTazz As Long
Private mTazzInfo As String

Public Sub Tazz(ar|g As Long, arg2 As string)
    mTazz = arg * 2
    mTazzInfo = arg2
End Sub
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, newName,
                    ("modOne", ComponentType.StandardModule, code));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [Test]
        [Category(nameof(ConflictDetectionSession))]
        public void RenameParameterInRecursiveSubroutine_HasConflict()
        {
            var selection = ("guzzArg", DeclarationType.Parameter);
            var inputCode =
$@"
Option Explicit

Private mGuzz As String
Private mCount As Long

Public Sub Guzz(guzzArg As String)
    mCount = mCount + 1
    mGuzz = guzzArg + CStr(mCount)
    If mCount < 6 Then Guzz mGuzz
End Sub
";
            Assert.AreEqual(true, TestForRenameConflict(inputCode, selection, "guzz"));
        }

        [TestCase("arg2", true)]
        [TestCase("arg3", false)]
        [TestCase("Tazz", true)]
        [TestCase("mTazz", true)] //Does not violate MS-VBAL conflict rules, but will change logic
        [Category(nameof(ConflictDetectionSession))]
        public void RenameFunctionParameter(string newName, bool isConflict)
        {
            var moduleContent1 =
$@"

Private mTazz As Long
Private mTazzInfo As String

Public Function Tazz(ar|g As Long, arg2 As string) As Long
    mTazz = arg * 2
    mTazzInfo = arg2
    Tazz = mTazz
End Function

Public Function Tazz2(arg As Long, arg3 As string) As Long
    mTazz = arg * 2
    mTazzInfo = arg3
    Tazz2 = mTazz
End Function
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, newName,
                    ("modOne", ComponentType.StandardModule, code));

            Assert.AreEqual(isConflict, hasConflict);
        }

        [TestCase("Bar1", DeclarationType.Function, "XYZ")]
        [TestCase("Bar2", DeclarationType.Procedure, "XYZ")]
        [TestCase("member11", DeclarationType.Variable, "XYZ")]
        [TestCase("member2", DeclarationType.Constant, "XYZ")]
        [Category(nameof(ConflictDetectionSession))]
        public void NoMatchingNames(string targetName, DeclarationType targetDeclarationType, string newName)
        {
            var selection = (targetName, targetDeclarationType);
            var content =
$@"
Private member11 As Long
Public Const member2 As Long = 5

Private Function Bar1() As Long
    Bar2
    Bar1 = member2 + member11
End Function

Private Sub Bar2()
    member2 = member2 * 4 
End Sub
";

            Assert.AreEqual(false, TestForRenameConflict(content, selection, newName));
        }

        [TestCase("Gazz", "Public", false)]
        [TestCase("mFazz", "Public", false)]
        [TestCase("EPvtTest", "Public", true)]
        [TestCase("EPvtTest", "Private", true)]
        [TestCase("TTestType", "Public", true)]
        [TestCase("TTestType", "Private", true)]
        [TestCase(MockVbeBuilder.TestModuleName, "Public", true)]
        [TestCase(MockVbeBuilder.TestModuleName, "Private", false)]
        [TestCase(MockVbeBuilder.TestProjectName, "Public", true)]
        [TestCase(MockVbeBuilder.TestProjectName, "Private", false)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void EnumerationRename_SameModuleConflicts(string newName, string accessibility, bool expected)
        {
            var inputCode =
$@"
Option Explicit

{accessibility} Enum ET|est
    FirstValue = 0
    SecondValue
End Enum

Private Enum EPvtTest
    ThirdValue = 10
    FourthValue
End Enum

Public Enum EPublicTest
    FifthValue = 20
    SixValue
End Enum

Private Type TTestType
    SeventhValue As Long
End Type

Private mFazz As ETest

Public Property Get Fazz() As ETest
    Fazz = mFazz
End Property

Public Sub Gazz(value As Long)
End Sub
";
            Assert.AreEqual(expected, TestForRenameConflict(newName, inputCode));
        }

        [Test]
        [Category(nameof(ConflictDetectionSession))]
        public void PublicEnumerationRenameToOtherProjectName_HasConflict()
        {
            var containingModuleContent =
$@"
Option Explicit

Public Enum ETest
    FirstValue = 0
    SecondValue
End Enum
";
            var conflictModuleCode =
$@"
Option Explicit
";
            var renameTargetModuleName = MockVbeBuilder.TestModuleName;
            var conflictProjectName = "ConflictProject";
            var conflictModuleName = "ConflictProjectModule";

            var projects = new Dictionary<string, IEnumerable<(string, string, ComponentType)>>()
            {
                [MockVbeBuilder.TestProjectName] = new List<(string, string, ComponentType)>() { (renameTargetModuleName, containingModuleContent, ComponentType.StandardModule) },
                [conflictProjectName] = new List<(string, string, ComponentType)>() { (conflictModuleName, conflictModuleCode, ComponentType.StandardModule), }
            };

            var builder = new MockVbeBuilder();
            foreach (var project in projects.Keys)
            {
                builder = AddProject(builder, project, projects[project]);
            }
            var vbe = builder.Build().Object;

            using (var parser = MockParser.CreateAndParse(vbe))
            {
                var target = parser.DeclarationFinder.UserDeclarations(DeclarationType.Enumeration)
                    .FirstOrDefault(item => item.IdentifierName.Equals("ETest"));

                var hasConflict = HasRenameConflict(parser, target, "ConflictProject");

                Assert.IsTrue(hasConflict);
            }
        }


        [TestCase("EConflictTest", true)]
        [TestCase("TConflictTest", true)]
        [TestCase("ConflictModule", true)]
        [TestCase("EPvtConflictTest", false)]
        [TestCase("TPvtConflictTest", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void PublicEnumerationRenamesConflictsOtherModule(string newName, bool expected)
        {
            var containingModuleContent =
$@"
Option Explicit

Public Enum ETest
    FirstValue = 0
    SecondValue
End Enum
";
            var conflictModuleCode =
$@"
Option Explicit

Public Enum EConflictTest
    FirstValue = 0
    SecondValue
End Enum

Public Type TConflictTest
    FirstValue As Long
    SecondValue As Long
End Type

Private Enum EPvtConflictTest
    FirstValuePvt = 0
    SecondValuePvt
End Enum

Private Type TPvtConflictTest
    FirstValuePvt As Long
    SecondValuePvt As Long
End Type
";
            var renameTargetModuleName = MockVbeBuilder.TestModuleName;
            var conflictModuleName = "ConflictModule";

            var vbe = MockVbeBuilder.BuildFromStdModules(
                        (renameTargetModuleName, containingModuleContent),
                        (conflictModuleName, conflictModuleCode));

            using (var parser = MockParser.CreateAndParse(vbe.Object))
            {
                var target = parser.DeclarationFinder.UserDeclarations(DeclarationType.Enumeration)
                    .FirstOrDefault(item => item.IdentifierName.Equals("ETest"));

                var hasConflict = HasRenameConflict(parser, target, newName);
                Assert.AreEqual(expected, hasConflict);
            }
        }

        [TestCase("Gazz", "Public", false)]
        [TestCase("mFazz", "Public", false)]
        [TestCase("EPvtTest", "Public", true)]
        [TestCase("EPvtTest", "Private", true)]
        [TestCase("TPublicTest", "Public", true)]
        [TestCase("TPublicTest", "Private", true)]
        [TestCase("TTestType", "Public", true)]
        [TestCase("TTestType", "Private", true)]
        [TestCase(MockVbeBuilder.TestModuleName, "Public", true)]
        [TestCase(MockVbeBuilder.TestModuleName, "Private", false)]
        [TestCase(MockVbeBuilder.TestProjectName, "Public", true)]
        [TestCase(MockVbeBuilder.TestProjectName, "Private", false)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void UDTRename_SameModuleConflicts(string newName, string accessibility, bool expected)
        {
            var inputCode =
$@"
Option Explicit

{accessibility} Type TT|est
    FirstValue As Long
End Type

Private Enum EPvtTest
    SecondValue = 10
End Enum

Public Type TPublicTest
    ThirdValue As Long
End Type

Private Type TTestType
    FourthValue As Long
End Type

Private mFazz As TTEst

Public Property Get Fazz() As TTEst
    Fazz = mFazz
End Property

Public Sub Gazz(value As TTEst)
End Sub
";
            Assert.AreEqual(expected, TestForRenameConflict(newName, inputCode));
        }

        [Test]
        [Category(nameof(ConflictDetectionSession))]
        public void PublicUDTRenameToOtherProjectName_HasConflict()
        {
            var containingModuleContent =
$@"
Option Explicit

Public Type TTest
    FirstValue As Long
End Type
";
            var conflictModuleCode =
$@"
Option Explicit
";
            var renameTargetModuleName = MockVbeBuilder.TestModuleName;
            var conflictProjectName = "ConflictProject";
            var conflictModuleName = "ConflictProjectModule";

            var projects = new Dictionary<string, IEnumerable<(string, string, ComponentType)>>()
            {
                [MockVbeBuilder.TestProjectName] = new List<(string, string, ComponentType)>() { (renameTargetModuleName, containingModuleContent, ComponentType.StandardModule) },
                [conflictProjectName] = new List<(string, string, ComponentType)>() { (conflictModuleName, conflictModuleCode, ComponentType.StandardModule), }
            };

            var builder = new MockVbeBuilder();
            foreach (var project in projects.Keys)
            {
                builder = AddProject(builder, project, projects[project]);
            }
            var vbe = builder.Build().Object;

            using (var parser = MockParser.CreateAndParse(vbe))
            {
                var target = parser.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedType)
                    .FirstOrDefault(item => item.IdentifierName.Equals("TTest"));

                var hasConflict = HasRenameConflict(parser, target, "ConflictProject");
                Assert.IsTrue(hasConflict);
            }
        }


        [TestCase("EConflictTest", true)]
        [TestCase("TConflictTest", true)]
        [TestCase("ConflictModule", true)]
        [TestCase("EPvtConflictTest", false)]
        [TestCase("TPvtConflictTest", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void PublicUDTRenamesConflictsOtherModule(string newName, bool expected)
        {
            var containingModuleContent =
$@"
Option Explicit

Public Type TTest
    FirstValue As Long
End Type
";
            var conflictModuleCode =
$@"
Option Explicit

Public Enum EConflictTest
    FirstValue = 0
End Enum

Public Type TConflictTest
    FirstValue As Long
End Type

Private Enum EPvtConflictTest
    FirstValuePvt = 0
End Enum

Private Type TPvtConflictTest
    FirstValuePvt As Long
End Type
";
            var renameTargetModuleName = MockVbeBuilder.TestModuleName;
            var conflictModuleName = "ConflictModule";

            var vbe = MockVbeBuilder.BuildFromStdModules(
                        (renameTargetModuleName, containingModuleContent),
                        (conflictModuleName, conflictModuleCode));

            using (var parser = MockParser.CreateAndParse(vbe.Object))
            {
                var target = parser.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedType)
                    .FirstOrDefault(item => item.IdentifierName.Equals("TTest"));

                var hasConflict = HasRenameConflict(parser, target, newName);
                Assert.AreEqual(expected, hasConflict);
            }
        }

        [TestCase("TTest")]
        [TestCase("ETest")]
        [Category(nameof(ConflictDetectionSession))]
        public void RenameProjectMatchesUDTOrEnum_HasConflict(string newName)
        {
            var sourceCode =
$@"
Option Explicit

Public Type TTest
    FirstValue As Long
End Type

Public Enum ETest
    FirstValue = 0
End Enum
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(sourceCode, out _);
            using (var parser = MockParser.CreateAndParse(vbe.Object))
            {
                var target = parser.DeclarationFinder.Projects.Single(p => p.IdentifierName.Equals(MockVbeBuilder.TestProjectName));

                var hasConflict = HasRenameConflict(parser, target, newName);
                Assert.IsTrue(hasConflict);
            }
        }

        [TestCase("Sub", "gOffset", true)]
        [TestCase("Sub", "gOffset2", false)]
        [TestCase("Function", "gOffset", true)]
        [TestCase("Function", "gOffset2", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void ClassModuleRenameMemberConflictWithNonQualifiedExternalReference(string memberType, string newName, bool expected)
        {
            var sourceCode =
$@"
Option Explicit

Private mValue As Long

Public {memberType} Tes|tThis(value As Long)
    mValue = gOffset + GlobalConstantModule.gOffset2 + value
End {memberType}
";

            var moduleCode =
$@"
Option Explicit

Public Const gOffset As Long = 100
Public Const gOffset2 As Long = 10
";
            (string code, Selection selection) = ToCodeAndSelectionTuple(sourceCode);

            var hasConflict = TestForRenameConflict(MockVbeBuilder.TestModuleName, selection, newName,
                    (MockVbeBuilder.TestModuleName, ComponentType.ClassModule, code),
                    ("GlobalConstantModule", ComponentType.StandardModule, moduleCode));

            Assert.AreEqual(expected, hasConflict);
        }

        [TestCase("Property Let", "gOffset", true)]
        [TestCase("Property Let", "gOffset2", false)]
        [Category(nameof(ConflictDetectionSession))]
        public void ClassModuleRenamePropertyConflictWithNonQualifiedExternalReference(string memberType, string newName, bool expected)
        {
            var sourceCode =
$@"
Option Explicit

Private mValue As Long

Public {memberType} Tes|tThis(value As Long)
    mValue = gOffset + GlobalConstantModule.gOffset2 + value
End Property
";

            var moduleCode =
$@"
Option Explicit

Public Const gOffset As Long = 100
Public Const gOffset2 As Long = 10
";
            (string code, Selection selection) = ToCodeAndSelectionTuple(sourceCode);

            var hasConflict = TestForRenameConflict(MockVbeBuilder.TestModuleName, selection, newName,
                    (MockVbeBuilder.TestModuleName, ComponentType.ClassModule, code),
                    ("GlobalConstantModule", ComponentType.StandardModule, moduleCode));

            Assert.AreEqual(expected, hasConflict);
        }

        [Test]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void RenameRespectsNewlyIntroducedFields()
        {
            var identifier = "TestFunc";
            var declarationType = DeclarationType.Function;

            var sourceCode =
$@"
Option Explicit

Private mTestVar As Long

Private Function TestFunc() As Long
End Function
";
            var newProxyVariables = new string[] { "ProxyVariable1", "ProxyVariable2", "ProxyVariable3" };

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(sourceCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(declarationType)
                                .Single(d => d.IdentifierName == identifier && d.QualifiedModuleName.ComponentName == MockVbeBuilder.TestModuleName);

                var factory = TestResolver.Resolve<IConflictSessionFactory>(state);

                var conflictSession = factory.Create();
                var renameConflictDetector = conflictSession.RenameConflictDetector;

                var fieldConflictDetector = conflictSession.NewEntityConflictDetector;

                var moduleProxy = conflictSession.ProxyCreator.Create(target.QualifiedModuleName);

                var nonConflictName = string.Empty;
                foreach (var newVariable in newProxyVariables)
                {
                    var proxy = conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, DeclarationType.Variable, newVariable, Accessibility.Private);
                    conflictSession.TryRegister(proxy, out _, true);
                }

                var targetProxy = conflictSession.ProxyCreator.Create(target, "ProxyVariable3");
                conflictSession.TryRegister(targetProxy, out _, true);
                StringAssert.AreEqualIgnoringCase("ProxyVariable4", (conflictSession.RenamePairs.Single()).NewName);
            }
        }

        private bool TestForRenameConflict(string inputCode, (string identifier, DeclarationType declarationType) selection, string newName)
        {
            var result = false;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var targets = state.DeclarationFinder.DeclarationsWithType(selection.declarationType);

                var target = targets.Single(d => d.IdentifierName == selection.identifier);

                result = HasRenameConflict(state, target, newName);
            }
            return result;
        }

        private bool TestForRenameConflict(string newName, string moduleContent)
        {
            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent);

            return TestForRenameConflict(MockVbeBuilder.TestModuleName, selection, newName, (MockVbeBuilder.TestModuleName, ComponentType.StandardModule, code));
        }

        private bool TestForRenameConflict(string selectionModuleName, Selection selection, string newName, params (string moduleName, ComponentType componentType, string inputCode)[] modules)
        {
            var builder = new MockVbeBuilder()
                            .ProjectBuilder(MockVbeBuilder.TestProjectName, ProjectProtection.Unprotected);

            foreach ((string moduleName, ComponentType componentType, string inputCode) in modules)
            {
                builder = builder.AddComponent(moduleName, componentType, inputCode);
            }

            var vbe = builder.AddProjectToVbeBuilder()
                            .Build();

            var result = false;
            var nonConflictName = string.Empty;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var module = state.DeclarationFinder.AllModules.First(qmn => qmn.ComponentName.Equals(selectionModuleName));
                var qualifiedSelection = new QualifiedSelection(module, selection);
                var target = state.DeclarationFinder.AllDeclarations
                                .Where(item => item.IsUserDefined)
                                .FirstOrDefault(item => item.IsSelected(qualifiedSelection) || item.References.Any(r => r.IsSelected(qualifiedSelection)));

                result = HasRenameConflict(state, target, newName);

            }
            return result;
        }

        private bool HasRenameConflict(RubberduckParserState state, Declaration target, string newName)
        {
            var renameConflictDetector = TestResolver.Resolve<IConflictSessionFactory>(state).Create().RenameConflictDetector;
            return renameConflictDetector.IsConflictingName(target, newName, out _);
        }

        private MockVbeBuilder AddProject(MockVbeBuilder builder, string projectName, IEnumerable<(string ModuleName, string Content, ComponentType ComponentType)> testComponents)
        {
            var enclosingProjectBuilder = builder.ProjectBuilder(projectName, ProjectProtection.Unprotected);

            foreach (var testComponent in testComponents)
            {
                enclosingProjectBuilder.AddComponent(testComponent.ModuleName, testComponent.ComponentType, testComponent.Content);
            }

            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            return builder;
        }

        private (string code, Selection selection) ToCodeAndSelectionTuple(string input)
        {
            var codeString = input.ToCodeString();
            return (codeString.Code, codeString.CaretPosition.ToOneBased());
        }
    }
}
