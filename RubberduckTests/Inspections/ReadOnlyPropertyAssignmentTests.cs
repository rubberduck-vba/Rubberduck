using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ReadOnlyPropertyAssignmentTests : InspectionTestsBase
    {
        [TestCase("MyData", 0)]
        [TestCase("MyData2", 1)] //Results in a readonly MyData property
        [Category("Inspections")]
        [Category("ReadOnlyPropertyAssignment")]
        public void SetUserDefinedClass_CalledFromSameModule(string setPropertyName, long expectedCount)
        {
            var sutInputCode =
$@"Private mData As AClass

Public Sub Test()
    Set MyData = New AClass
End Sub

Public Property Get MyData() As AClass
    Set MyData = mData
End Property

Public Property Set {setPropertyName}(RHS As AClass)
    Set mData = MyData
End Property
";
            Assert.AreEqual(expectedCount, InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, sutInputCode, ComponentType.StandardModule),
                ("AClass", $"Option Explicit{Environment.NewLine}", ComponentType.ClassModule)).Count());
        }

        [TestCase("TheVariant", 0)]
        [TestCase("TheVariant2", 1)] //Results in a readonly MyData property
        [Category("Inspections")]
        [Category("ReadOnlyPropertyAssignment")]
        public void VariantLet_CalledFromSameModule(string letPropertyName, long expectedCount)
        {
            var sutInputCode =
$@"Option Explicit

Private myVariant As Variant

Public Sub Test()
    TheVariant = 7
End Sub

Public Property Get TheVariant() As Variant
    If IsObject(myVariant) Then
        Set TheVariant = myVariant
    Else
        TheVariant = myVariant
    End If
End Property
Public Property Let {letPropertyName}(RHS As Variant)
    myVariant = RHS
End Property
";
            Assert.AreEqual(expectedCount, InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, sutInputCode, ComponentType.StandardModule)).Count());
        }

        [TestCase("TheVariant", 0)]
        [TestCase("TheVariant2", 1)] //Results in a readonly MyData property
        [Category("Inspections")]
        [Category("ReadOnlyPropertyAssignment")]
        public void VariantSet_CalledFromSameModule(string setPropertyName, long expectedCount)
        {
            var sutInputCode =
$@"Option Explicit

Private myVariant As Variant

Public Sub Test()
    Set TheVariant = new AClass
End Sub

Public Property Get TheVariant() As Variant
    If IsObject(myVariant) Then
        Set TheVariant = myVariant
    Else
        TheVariant = myVariant
    End If
End Property
Public Property Set {setPropertyName}(RHS As Variant)
    Set myVariant = RHS
End Property
";

            Assert.AreEqual(expectedCount, InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, sutInputCode, ComponentType.StandardModule),
                ("AClass", $"Option Explicit{Environment.NewLine}", ComponentType.ClassModule)).Count());
        }

        [TestCase("MyData", 0)]
        [TestCase("MyData2", 1)] //Results in a readonly MyData property
        [Category("Inspections")]
        [Category("ReadOnlyPropertyAssignment")]
        public void ObjectDataType_CalledFromOtherModule(string setPropertyName, long expectedCount)
        {
            var sutInputCode =
$@"
Option Explicit

Private mData As Collection

Public Sub Test()
    Set MyData = New Collection
End Sub

Public Property Get MyData() As Collection
    Set MyData = mData
End Property

Public Property Set {setPropertyName}(ByVal RHS As Collection)
    Set mData = RHS
End Property
";

            var callingModule =
$@"
Option Explicit

Public Sub Test()
    Set {MockVbeBuilder.TestModuleName}.MyData = New Collection
End Sub
";

            Assert.AreEqual(expectedCount, InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, sutInputCode, ComponentType.StandardModule),
                ("CallingModule", callingModule, ComponentType.StandardModule)).Count());
        }

        [TestCase("MyData", 0)]
        [TestCase("MyData2", 1)] //Results in a readonly MyData property
        [Category("Inspections")]
        [Category("ReadOnlyPropertyAssignment")]
        public void ScalarDataType_CalledFromOtherModule(string letPropertyName, long expectedCount)
        {
            var sutInputCode =
$@"
Option Explicit

Private mData As Long

Public Sub Test()
    Set MyData = 8
End Sub

Public Property Get MyData() As Long
    MyData = mData
End Property

Public Property Let {letPropertyName}(ByVal RHS As Long)
    mData = RHS
End Property
";

            var callingModule =
$@"
Option Explicit

Public Sub Test()
    {MockVbeBuilder.TestModuleName}.MyData = 80
End Sub
";

            Assert.AreEqual(expectedCount, InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, sutInputCode, ComponentType.StandardModule),
                ("CallingModule", callingModule, ComponentType.StandardModule)).Count());
        }

        //TODO: Remove 'Ignore' once this false positive scenario is resolved
        [Test]
        [Ignore("False Positive")]
        [Category("Inspections")]
        [Category("ReadOnlyPropertyAssignment")]
        public void FalsePositiveMCVE_ShouldNotFlag()
        {
            var fooClassCode =
@"
Option Explicit

Private mFooValue As String

'If you remove this Sub, the false positive does not occur
Private Sub Class_Initialize()
    mFooValue = ""Test""
End Sub

Public Property Get FooValue() As String
    FooValue = mFooValue
End Property

Public Property Let FooValue(ByVal RHS As String)
    mFooValue = RHS
End Property
";
            var sutInputCode =
@"
Option Explicit

'If Sub Test is placed after Property Get FooValue the
'False Positive does not occur
Public Sub Test()
    Dim fc As FooClass
    Set fc = New FooClass
    fc.FooValue = FooValue
End Sub

Public Property Get FooValue() As String
    FooValue = ""Test""
End Property
";

            Assert.AreEqual(0, InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, sutInputCode, ComponentType.StandardModule),
                ("FooClass", fooClassCode, ComponentType.ClassModule)).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ReadOnlyPropertyAssignmentInspection(state);
        }
    }
}
