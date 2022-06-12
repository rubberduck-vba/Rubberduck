using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using NUnit.Framework;
using System;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using RubberduckTests.Mocks;
using System.Threading;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.Grammar;

namespace RubberduckTests.Inspections
{
    class PublicEnumerationDeclaredInWorksheetInspectionTests
    {
        private static string[] _worksheetSuperTypeNames = new string[] { "Worksheet", "_Worksheet" };

        [Test]
        [Category("Inspections")]
        [Category(nameof(PublicEnumerationDeclaredInWorksheetInspection))]
        public void EnumerationDeclaredWithinWorksheet_InspectionName()
        {
            var inspection = new PublicEnumerationDeclaredInWorksheetInspection(null);

            Assert.AreEqual(nameof(PublicEnumerationDeclaredInWorksheetInspection), inspection.Name);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(PublicEnumerationDeclaredInWorksheetInspection))]
        [TestCase(ComponentType.StandardModule)]
        [TestCase(ComponentType.ClassModule)]
        [TestCase(ComponentType.UserForm)]
        public void NonDocumentComponentsAreNotFlagged(ComponentType componentType)
        {

            var code = 
@"Option Explicit

Public Enum DeclaredEnum
    wsMember1 = 0
    wsMember1 = 1
End Enum
";
            var module = new DocumentModuleFake(code, new string[] { });

            var inspectionResults = InspectionResults(module);
            int actual = inspectionResults.Count();

            Assert.AreEqual(0, actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(PublicEnumerationDeclaredInWorksheetInspection))]
        [TestCase("Public ", 1)]
        [TestCase("Private ", 0)]
        [TestCase("", 1)]
        public void FlagsPublicEnumerationsOnly(string accessibility, int expected)
        {

            var code = 
$@"Option Explicit

{accessibility}Enum DeclaredEnum
    wsMember1 = 0
    wsMember1 = 1
End Enum
";
            var docModuleStub = new DocumentModuleFake(code, _worksheetSuperTypeNames);
            var inspectionResults = InspectionResults(docModuleStub);
            int actual = inspectionResults.Count();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(PublicEnumerationDeclaredInWorksheetInspection))]
        [TestCase(true, 1)]
        [TestCase(false, 0)]
        public void FlagsWorksheetDocumentTypesOnly(bool isWorksheetDoc, int expected)
        {
            var code = 
$@"Option Explicit

Public Enum DeclaredEnum
    wsMember1 = 0
    wsMember1 = 1
End Enum
";

            var docModuleStub = isWorksheetDoc
                ? new DocumentModuleFake(code, _worksheetSuperTypeNames)
                : new DocumentModuleFake(code, new string[] { "NotAWorksheetDocument" });

            var inspectionResults = InspectionResults(docModuleStub);
            int actual = inspectionResults.Count();

            Assert.AreEqual(expected, actual);
        }

        [TestCase("_Worksheet")]
        [TestCase("Worksheet")]
        [Category("Inspections")]
        [Category(nameof(PublicEnumerationDeclaredInWorksheetInspection))]
        public void WorksheetDocument_FlagsOnEachWorksheetSuperTypeName(string superTypeName)
        {
            var code = 
$@"Option Explicit

Public Enum DeclaredEnum
    wsMember1 = 0
    wsMember1 = 1
End Enum
";

            var docModuleStub = new DocumentModuleFake(code, new string[] { superTypeName });
            var inspectionResults = InspectionResults(docModuleStub);
            int actual = inspectionResults.Count();

            Assert.AreEqual(1, actual);
        }

        private IEnumerable<IInspectionResult> InspectionResults(DocumentModuleFake docModule)
        {
            var vbe = MockVbeBuilder.BuildFromModules(docModule.AsTestModuleTuple).Object;
            return GetInspectionResults(vbe, docModule);
        }

        private IEnumerable<IInspectionResult> GetInspectionResults(IVBE vbe, DocumentModuleFake testDocument)
        {
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var inspection = new PublicEnumerationDeclaredInWorksheetInspection(state);

                if (testDocument.SuperTypeNames.Any())
                {
                    inspection.RetrieveSuperTypeNames = new Func<ClassModuleDeclaration, IEnumerable<string>>((d) => testDocument.SuperTypeNames);
                }

                return inspection.GetInspectionResults(CancellationToken.None);
            }
        }

        /// <summary>
        /// Wraps a Document module to support faking the ClassModuleDeclaration's SuperTypeNames property
        /// </summary>
        private class DocumentModuleFake
        {
            private string[] _superTypeNames;

            private string _code;

            public DocumentModuleFake(string code, string[] superTypeNames)
            {
                _code = code;
                _superTypeNames = superTypeNames;
            }

            public (string, string, ComponentType) AsTestModuleTuple => ("TestDocumentModule", _code, ComponentType.Document);

            public IEnumerable<string> SuperTypeNames => _superTypeNames;
        }
    }
}
