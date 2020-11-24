using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using NUnit.Framework;

namespace RubberduckTests.Inspections
{
    class EnumerationDeclaredWithinWorksheetInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void EnumerationDeclaredWithinWorksheet_InspectionName()
        {
            var inspection = new EnumerationDeclaredWithinWorksheetInspection(null, null);

            Assert.AreEqual(nameof(EnumerationDeclaredWithinWorksheetInspection), inspection.Name);
        }

        [Test]
        [Category("Inspections")]
        [Category("EnumerationDeclaredWithinWorksheet")]
        public void Project_with_multiple_enumerations_flags_only_enum_declared_within_worksheets()
        {
            #region InputCode
            const string worksheetCode = @"Option Explicit
Public Enum WorksheetEnum
    wsMember1 = 0
    wsMember1 = 1
End Enum";
            const string standardModuleCode = @"Option Explicit
Public Enum StdModEnum
    stdMember1 = 0
    stdMember1 = 2
End Enum";
            const string classModuleCode = @"Option Explicit
Public Enum ClassModEnum
    classMember1 = 0
    classMember2 = 3
End Enum";
            const string userFormModuleCode = @"Option Explicit
Public Enum FormEnum
    formMember1 = 0
    formMember2 = 4
End Enum";
            #endregion

            var vbeBuilder = new MockVbeBuilder();
            var project = vbeBuilder.ProjectBuilder("Project1", ProjectProtection.Unprotected)
                .AddComponent("Sheet", ComponentType.DocObject, worksheetCode)
                .AddComponent("StdModule", ComponentType.StandardModule, standardModuleCode)
                .AddComponent("ClsMod", ComponentType.ClassModule, classModuleCode)    
                .AddComponent("UserFormMod", ComponentType.UserForm, userFormModuleCode)
                .Build();

            vbeBuilder.AddProject(project);
            
            var vbe = vbeBuilder.Build();

            var inspectionResults = InspectionResults(vbe.Object);
            int actual = inspectionResults.Count();

            Assert.AreEqual(1, actual);
        }

        [Test]
        [Category("Inspections")]
        [Category("EnumerationDeclaredWithinWorksheet")]
        [TestCase(ComponentType.ActiveXDesigner)]
        [TestCase(ComponentType.ClassModule)]
        [TestCase(ComponentType.ComComponent)]
        [TestCase(ComponentType.Document)]
        [TestCase(ComponentType.MDIForm)]
        [TestCase(ComponentType.PropPage)]
        [TestCase(ComponentType.RelatedDocument)]
        [TestCase(ComponentType.ResFile)]
        [TestCase(ComponentType.StandardModule)]
        [TestCase(ComponentType.Undefined)]
        [TestCase(ComponentType.UserControl)]
        [TestCase(ComponentType.UserForm)]
        [TestCase(ComponentType.VBForm)]
        public void Enumerations_declared_within_non_worksheet_object_have_no_inpsection_result(ComponentType componentType)
        {
            const string code = @"Option Explicit
Public Enum TestEnum
    Member1
    Member2
End Enum";

            var vbeBuilder = new MockVbeBuilder();
            var project = vbeBuilder.ProjectBuilder("UnitTestProject", ProjectProtection.Unprotected)
                .AddComponent(componentType.ToString() + "module", componentType, code)
                .Build();

            vbeBuilder.AddProject(project);

            var vbe = vbeBuilder.Build();

            var inspectionResults = InspectionResults(vbe.Object);
            int actual = inspectionResults.Count();

            Assert.AreEqual(0, actual);
        }

        [Test]
        [Category("Inspections")]
        [Category("EnumerationDeclaredWithinWorksheet")]
        public void Private_type_declared_within_worksheet_has_no_inspection_result()
        {
            const string code = @"Option Explicit

Private Type THelper
    Name As String
    Address As String
End Type

Private this as THelper";

            var vbeBuilder = new MockVbeBuilder();
            var project = vbeBuilder.ProjectBuilder("UnitTestProject", ProjectProtection.Unprotected)
                .AddComponent("WorksheetForTest", ComponentType.DocObject, code)
                .Build();

            var vbe = vbeBuilder.Build();

            var inspectionResults = InspectionResults(vbe.Object);
            int actual = inspectionResults.Count();

            Assert.AreEqual(0, actual);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EnumerationDeclaredWithinWorksheetInspection(state, state.ProjectsProvider);
        }
    }
}
