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
        [Test]
        [Category("Inspections")]
        [Category("ReadOnlyPropertyAssignment")]
        public void SetUserDefinedClassMCVE_Flags()
        {
            var sutInputCode =
@"Private mData As AClass

Public Sub Test()
    Set MyData = New AClass
End Sub

Public Property Get MyData() As AClass
    Set MyData = mData
End Property
";
            Assert.AreEqual(1, InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, sutInputCode, ComponentType.StandardModule),
                ("AClass", $"Option Explicit{Environment.NewLine}", ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("ReadOnlyPropertyAssignment")]
        public void SetUserDefinedClassSetExists_NotFlagged()
        {
            var sutInputCode =
@"Private mData As AClass

Public Sub Test()
    Set MyData = New AClass
End Sub

Public Property Get MyData() As AClass
    Set MyData = mData
End Property
Public Property Set MyData(RHS As AClass)
    Set mData = MyData
End Property
";
            Assert.AreEqual(0, InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, sutInputCode, ComponentType.StandardModule),
                ("AClass", $"Option Explicit{Environment.NewLine}", ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("ReadOnlyPropertyAssignment")]
        public void LetVariantMCVE_Flags()
        {
            var sutInputCode =
@"Option Explicit

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
";

            Assert.AreEqual(1, InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, sutInputCode, ComponentType.StandardModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("ReadOnlyPropertyAssignment")]
        public void LetVariantLetExists_NotFlagged()
        {
            var sutInputCode =
@"Option Explicit

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
Public Property Let TheVariant(RHS As Variant)
    myVariant = RHS
End Property
";
            Assert.AreEqual(0, InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, sutInputCode, ComponentType.StandardModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("ReadOnlyPropertyAssignment")]
        public void SetVariantMCVE_Flags()
        {
            var sutInputCode =
@"Option Explicit

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
";

            Assert.AreEqual(1, InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, sutInputCode, ComponentType.StandardModule),
                ("AClass", $"Option Explicit{Environment.NewLine}", ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("ReadOnlyPropertyAssignment")]
        public void SetVariantSetExists_NotFlagged()
        {
            var sutInputCode =
@"Option Explicit

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
Public Property Set TheVariant(RHS As Variant)
    Set myVariant = RHS
End Property
";

            Assert.AreEqual(0, InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, sutInputCode, ComponentType.StandardModule),
                ("AClass", $"Option Explicit{Environment.NewLine}", ComponentType.ClassModule)).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ReadOnlyPropertyAssignmentInspection(state);
        }
    }
}
