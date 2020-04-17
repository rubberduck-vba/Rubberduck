using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using TestDI = RubberduckTests.Refactoring.NameConflictFinderTestsDI;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class NameConflictFinderRenameTests
    {
        //MS_VBAL 5.4.3.1, 5.4.3.2
        //Method names need to be different than contained local variables or constants
        [TestCase("Dim goo As Long", "GOO", true)]
        [TestCase("Static goo As Long", "GOO", true)]
        [TestCase("Const goo As Long = 10", "GOO", true)]
        [Category("Refactoring")]
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        //(Speculation) If VBA treats parameter references the same as local variable references 
        //within the procedure, then
        //MS-VBAL 5.4.3.1 (Method names need to be different than contained local variables)
        //would be the applicable conflict 'rule'.
        [Test]
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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

        [TestCase("member1", true)]
        [TestCase("adder", false)]
        [TestCase("Foo", true)]
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [TestCase("Gazz", true)]
        [TestCase("Buzz", true)]
        [TestCase("Fooz", true)]
        [TestCase("Tizz", true)]
        [TestCase("Bazz2", false)]
        [Category(nameof(NameConflictFinder))]
        public void RenamedSubHasExternalNonQualifiedReferenceConflicts(string nameToCheck, bool isConflict)
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
        [Category(nameof(NameConflictFinder))]
        public void RenamedSubConflictsWithParameter(string nameToCheck, bool isConflict)
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
        [Category(nameof(NameConflictFinder))]
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

        [Category(nameof(NameConflictFinder))]
        public void PrivateSub_CheckConflictsInOtherModule()
        {
            var nameToCheck = "DoThis";
            var moduleContent1 =
@"
Private Sub DoThis(filename As String)
    SetFi|lename filename
End Sub
";
            var moduleContent2 =
@"
Private member1 As String

Public Sub SetFilename(filename As String)
    member1 = filename
End Sub
";

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modTwo", ComponentType.StandardModule, moduleContent2));

            Assert.AreEqual(true, hasConflict);
        }

        [TestCase("DoThis", true)]
        [TestCase("filename", true)]
        [TestCase("member1", true)]
        [TestCase("mFolderpath", true)]
        [TestCase("ExtractFilename", true)]
        [TestCase("StoreFilename", true)]
        [TestCase("filepath", true)]
        [Category(nameof(NameConflictFinder))]
        public void PrivateSub_MultipleReferences(string nameToCheck, bool isConflict)
        {

            var moduleContent1 =
@"
Private Sub DoThis(filename As String)
    SetFil|ename filename
End Sub
";
            var moduleContent2 =
@"
Private member1 As String

Public Sub SetFilename(filename As String)
    member1 = filename
End Sub
";
            var moduleContent3 =
@"
Private mFolderpath As String

Private Sub StoreFilename(filepath As String)
    Dim filename As String
    filename = ExtractFilename(filepath)
    SetFilename filename
End Sub

Private Function ExtractFilename(filepath As String) As String
    ExtractFilename = filepath
End Function"
;

            (string code, Selection selection) = ToCodeAndSelectionTuple(moduleContent1);

            var hasConflict = TestForRenameConflict("modOne", selection, nameToCheck,
                    ("modOne", ComponentType.StandardModule, code),
                    ("modTwo", ComponentType.StandardModule, moduleContent2),
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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

                var conflictFinder = TestDI.Resolve<INameConflictFinder>(parser);
                var hasConflict = conflictFinder.RenameCreatesNameConflict(target, projectTwoModule, out _);
                Assert.IsFalse(hasConflict);
            }
        }

        [Test]
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [TestCase("Tazz", true)]
        [TestCase("mTazz", true)] //Does not violate MS-VBAL conflict rules, but will change logic
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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

                var conflictFinder = TestDI.Resolve<INameConflictFinder>(parser);
                var hasConflict = conflictFinder.RenameCreatesNameConflict(target, "ConflictProject", out _);
                Assert.IsTrue(hasConflict);
            }
        }


        [TestCase("EConflictTest", true)]
        [TestCase("TConflictTest", true)]
        [TestCase("ConflictModule", true)]
        [TestCase("EPvtConflictTest", false)]
        [TestCase("TPvtConflictTest", false)]
        [Category(nameof(NameConflictFinder))]
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

                var conflictFinder = TestDI.Resolve<INameConflictFinder>(parser);
                var hasConflict = conflictFinder.RenameCreatesNameConflict(target, newName, out _);
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
        [Category(nameof(NameConflictFinder))]
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
        [Category(nameof(NameConflictFinder))]
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

                var conflictFinder = TestDI.Resolve<INameConflictFinder>(parser);
                var hasConflict = conflictFinder.RenameCreatesNameConflict(target, "ConflictProject", out _);
                Assert.IsTrue(hasConflict);
            }
        }


        [TestCase("EConflictTest", true)]
        [TestCase("TConflictTest", true)]
        [TestCase("ConflictModule", true)]
        [TestCase("EPvtConflictTest", false)]
        [TestCase("TPvtConflictTest", false)]
        [Category(nameof(NameConflictFinder))]
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

                var conflictFinder = TestDI.Resolve<INameConflictFinder>(parser);
                var hasConflict = conflictFinder.RenameCreatesNameConflict(target, newName, out _);
                Assert.AreEqual(expected, hasConflict);
            }
        }

        [TestCase("TTest")]
        [TestCase("ETest")]
        [Category(nameof(NameConflictFinder))]
        public void RenameProjectMatchesUDTOrEnum_HasConflict(string newName)
        {
            var containingModuleContent =
$@"
Option Explicit

Public Type TTest
    FirstValue As Long
End Type

Public Enum ETest
    FirstValue = 0
End Enum
";
            var conflictModuleCode =
$@"
Option Explicit
";
            var renameTargetModuleName = MockVbeBuilder.TestModuleName;
            var otherProject = "ConflictProject";
            var conflictModuleName = "ConflictProjectModule";

            var projects = new Dictionary<string, IEnumerable<(string, string, ComponentType)>>()
            {
                [MockVbeBuilder.TestProjectName] = new List<(string, string, ComponentType)>() { (renameTargetModuleName, containingModuleContent, ComponentType.StandardModule) },
                [otherProject] = new List<(string, string, ComponentType)>() { (conflictModuleName, conflictModuleCode, ComponentType.StandardModule), }
            };

            var builder = new MockVbeBuilder();
            foreach (var project in projects.Keys)
            {
                builder = AddProject(builder, project, projects[project]);
            }
            var vbe = builder.Build().Object;

            using (var parser = MockParser.CreateAndParse(vbe))
            {
                var target = parser.DeclarationFinder.Projects.Single(p => p.IdentifierName.Equals(otherProject));

                var conflictFinder = TestDI.Resolve<INameConflictFinder>(parser);
                var hasConflict = conflictFinder.RenameCreatesNameConflict(target, newName, out _);
                Assert.IsTrue(hasConflict);
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

                var conflictFinder = TestDI.Resolve<INameConflictFinder>(state);
                result = conflictFinder.RenameCreatesNameConflict(target, newName, out _);
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var module = state.DeclarationFinder.AllModules.First(qmn => qmn.ComponentName.Equals(selectionModuleName));
                var qualifiedSelection = new QualifiedSelection(module, selection);
                var target = state.DeclarationFinder.AllDeclarations
                                .Where(item => item.IsUserDefined)
                                .FirstOrDefault(item => item.IsSelected(qualifiedSelection) || item.References.Any(r => r.IsSelected(qualifiedSelection)));


                var conflictFinder = TestDI.Resolve<INameConflictFinder>(state);
                result = conflictFinder.RenameCreatesNameConflict(target, newName, out _);
            }
            return result;
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
