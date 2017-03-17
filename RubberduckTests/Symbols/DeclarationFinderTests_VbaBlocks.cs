using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Symbols
{
    public static class DeclarationFinderTests_VbaBlocks
    {
        public static string InProcedure_MethodDeclaration_moduleContent1()
        {
            return
    @"

Private member1 As Long

Public Function Foo() As Long   'Selecting 'Foo' to rename
    Dim adder as Long
    adder = 10
    member1 = member1 + adder
    Foo = member1
End Function
";
        }
        public static string InProcedure_LocalVariableReference_moduleContent1()
        {
            return
    @"

Private member1 As Long

Public Function Foo() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder   'Selecting 'adder' to rename
    Foo = member1
End Function
";
        }

        public static string InProcedure_MemberDeclaration_moduleContent1()
        {
            return
    @"

Private member1 As Long

Public Function Foo() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder       'Selecting member1 to rename
    Foo = member1
End Function
";
        }

        public static string ModuleScope_CFirstClassContent()
        {
            return
    @"

Private member1 As Long

Public Function Foo() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder       'Selecting 'member1' to rename
    Foo = member1
End Function
";
        }

        public static string ModuleScope_moduleContent2()
        {
            return
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
    member2 = member2 * 4
End Sub
";
        }

        public static string PublicClassAndPubicModuleSub_CFirstClass()
        {
            return
    @"
Public Function Foo() As Long   'Selecting 'Foo' to rename
    Foo = 5
End Function
";
        }

        public static string PublicClassAndPubicModuleSub_moduleContent2()
        {
            return
    @"
Public Function Foo2() As Long
    Foo2 = 2
End Function
";
        }

        public static string Module_To_ClassScope_moduleContent1()
        {
            return
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
    member2 = member2 * 4   'Selecting member2 to rename
End Sub
";
        }

        public static string Module_To_ClassScope_CFirstClass()
        {
            return
    @"

Private member1 As Long

Public Function Foo() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder
    Foo = member1
End Function
";
        }

        public static string PrivateSub_RespectPublicSubInOtherModule_moduleContent1()
        {
            return
@"


Private Sub DoThis(filename As String)
    SetFilename filename            'Selecting 'SetFilename' to rename
End Sub
";
        }

        public static string PrivateSub_RespectPublicSubInOtherModule_moduleContent2()
        {
            return
    @"

Private member1 As String

Public Sub SetFilename(filename As String)
    member1 = filename
End Sub
";
        }

        public static string PrivateSub_MultipleReferences_moduleContent1()
        {
            return
@"


Private Sub DoThis(filename As String)
    SetFilename filename       'Selecting 'SetFilename' to rename
End Sub
";
        }

        public static string PrivateSub_MultipleReferences_moduleContent2()
        {
            return
    @"

Private member1 As String

Public Sub SetFilename(filename As String)
    member1 = filename
End Sub
";
        }

        public static string PrivateSub_MultipleReferences_moduleContent3()
        {
            return
    @"

Private mFolderpath As String

Private Sub StoreFilename(filepath As String)
    Dim filename As String
    filename = ExtractFilename(filepath)
    SetFilename filename
End Sub

Private Function ExtractFilename(filepath As String) As String
    ExtractFilename = filepath
End Function
";
        }

        public static string PrivateSub_WithBlock_CFileHelperContent()
        {
            return
    @"

Private mFolderpath As String

Public Sub StoreFilename(input As String)
    Dim filename As String
    filename = ExtractFilename(input)
    SetFilename filename
End Sub

Private Function ExtractFilename(filepath As String) As String
    ExtractFilename = filepath
End Function

Public Sub Bar()
End Sub
";
        }

        public static string PrivateSub_WithBlock_ModuleContent1()
        {
            return
    @"

Private myData As String
Private mDupData As String

Public Sub Foo(filenm As String)
    Dim filepath As String
    filepath = ""C:\MyStuff\"" & filenm
    Dim helper As CFileHelper
    Set helper = new CFileHelper
    With helper
        .StoreFilename filepath     'Selecting 'StoreFilename' to rename
        mDupData = filepath
    End With
End Sub

Public Sub StoreFilename(filename As String)
    myData = filename
End Sub
";
        }

        public static string Module_To_ModuleScopeResolution__moduleContent1()
        {
            return
    @"

Private member11 As Long
Public member2 As Long

Private Function Bar1() As Long
    Bar2
    Bar1 = member2 + modTwo.Foo1 + modTwo.Foo2 + modTwo.Foo3   'Selecting Foo2 to rename
End Function

Private Sub Bar2()
    member2 = member2 * 4 
End Sub
";
        }
        public static string Module_To_ModuleScopeResolution__moduleContent2()
        {
            return
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
        }

        public static string FiendishlyAmbiguousNameSelectsSmallestScopedDeclaration()
        {
            return @"
Option Explicit

Public Sub foo()
    Dim foo As Long
    foo = 42
    Debug.Print foo
End Sub
";
        }

        public static string AmbiguousNameSelectsSmallestScopedDeclaration()
        {
            return @"
Option Explicit

Public Sub foo()
    Dim foo As Long
    foo = 42
    Debug.Print foo
End Sub
";
        }
    }
}
