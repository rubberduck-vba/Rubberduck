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
        public void Project_with_multiple_enumerations_only_returns_worksheet_result()
        {
            #region InputCode
            const string worksheet1Code = @"Option Explicit
Public Enum Worksheet1Enumeration
    Option1 = 0
    Option2 = 1
    Option3 = 3
    Option4 = 7
End Enum
";
            const string worksheet2Code = @"Option Explicit
Dim bar As Long";
            const string standardModuleCode = @"Option Explicit
Public Enum StdModEnumeration
    Option1 = 0
    Option2 = 1
    Option3 = 2
    Option4 = 4
End Enum";
            const string classModuleCode = @"Option Explicit
Public Enum ClassModEnum
    Option1 = 0
    Option2 = 1
End Enum";
            #endregion

            var vbeBuilder = new MockVbeBuilder();
            var project = vbeBuilder.ProjectBuilder("Project1", ProjectProtection.Unprotected)
                .AddComponent("Sheet1", ComponentType.DocObject, worksheet1Code)
                .AddComponent("Sheet2", ComponentType.DocObject, worksheet2Code)
                .AddComponent("Module1", ComponentType.StandardModule, standardModuleCode)
                .AddComponent("ClsMod", ComponentType.ClassModule, classModuleCode)                
                .Build();

            vbeBuilder.AddProject(project);
            
            var vbe = vbeBuilder.Build();

            int actual = InspectionResults(vbe.Object).Count();

            Assert.AreEqual(1, actual);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EnumerationDeclaredWithinWorksheetInspection(state, state.ProjectsProvider);
        }
    }
}
