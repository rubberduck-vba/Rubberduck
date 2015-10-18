using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// Defines <see cref="Declaration"/> objects for the Excel object model, loaded when host application is Microsoft Excel.
    /// </summary>
    internal static class ExcelObjectModel
    {
        private static IEnumerable<Declaration> _excelDeclarations;
        private static readonly QualifiedModuleName ExcelModuleName = new QualifiedModuleName("Excel", "Excel");

        public static IEnumerable<Declaration> Declarations
        {
            get
            {
                if (_excelDeclarations == null)
                {
                    var nestedTypes = typeof(ExcelObjectModel).GetNestedTypes(BindingFlags.NonPublic);
                    var fields = nestedTypes.SelectMany(t => t.GetFields());
                    var values = fields.Select(f => f.GetValue(null));
                    _excelDeclarations = values.Cast<Declaration>();
                }

                return _excelDeclarations;
            }
        }

        private class ExcelLib
        {
            public static readonly Declaration Excel = new Declaration(new QualifiedMemberName(ExcelModuleName, "Excel"), null, "Excel", "Excel", true, false, Accessibility.Global, DeclarationType.Project);

            private static readonly QualifiedModuleName RangeModuleName = new QualifiedModuleName("Excel", "Range");
            public static readonly Declaration Range = new Declaration(new QualifiedMemberName(RangeModuleName, "Range"), ExcelLib.Excel, "Excel", "Range", false, false, Accessibility.Global, DeclarationType.Class);

            public static Declaration Activate = new Declaration(new QualifiedMemberName(RangeModuleName, "Activate"), Range, "Excel.Range", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Address = new Declaration(new QualifiedMemberName(RangeModuleName, "Address"), Range, "Excel.Range", "String", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Select = new Declaration(new QualifiedMemberName(RangeModuleName, "Select"), Range, "Excel.Range", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Cells = new Declaration(new QualifiedMemberName(RangeModuleName, "Cells"), Range, "Excel.Range", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Formula = new Declaration(new QualifiedMemberName(RangeModuleName, "Formula"), Range, "Excel.Range", "Variant", false, false, Accessibility.Global, DeclarationType.PropertyGet);
        }

        private class GlobalClass
        {
            public static readonly Declaration Global = new Declaration(new QualifiedMemberName(ExcelModuleName, "Global"), ExcelLib.Excel, "Excel", "Global", false, false, Accessibility.Global, DeclarationType.Class);

            public static Declaration Evaluate = new Declaration(new QualifiedMemberName(ExcelModuleName, "Evaluate"), Global, "Excel.Global", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Range = new Declaration(new QualifiedMemberName(ExcelModuleName, "Range"), Global, "Excel.Global", "Range", true, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration Selection = new Declaration(new QualifiedMemberName(ExcelModuleName, "Selection"), Global, "Excel.Global", "Object", true, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static Declaration Activate = new Declaration(new QualifiedMemberName(ExcelModuleName, "Activate"), Global, "Excel.Global", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Select = new Declaration(new QualifiedMemberName(ExcelModuleName, "Select"), Global, "Excel.Global", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Cells = new Declaration(new QualifiedMemberName(ExcelModuleName, "Cells"), Global, "Excel.Global", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Names = new Declaration(new QualifiedMemberName(ExcelModuleName, "Names"), Global, "Excel.Global", "Names", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Sheets = new Declaration(new QualifiedMemberName(ExcelModuleName, "Sheets"), Global, "Excel.Global", "Sheets", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Worksheets = new Declaration(new QualifiedMemberName(ExcelModuleName, "Worksheets"), Global, "Excel.Global", "Worksheets", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration WorksheetFunction = new Declaration(new QualifiedMemberName(ExcelModuleName, "WorksheetFunction"), Global, "Excel.Global", "WorksheetFunction", false, false, Accessibility.Global, DeclarationType.PropertyGet);

            public static Declaration Columns = new Declaration(new QualifiedMemberName(ExcelModuleName, "Columns"), Global, "Excel.Global", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Rows = new Declaration(new QualifiedMemberName(ExcelModuleName, "Rows"), Global, "Excel.Global", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);

            public static Declaration ActiveCell = new Declaration(new QualifiedMemberName(ExcelModuleName, "ActiveCell"), Global, "Excel.Global", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration ActiveSheet = new Declaration(new QualifiedMemberName(ExcelModuleName, "ActiveSheet"), Global, "Excel.Global", "Object", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration ActiveWorkbook = new Declaration(new QualifiedMemberName(ExcelModuleName, "ActiveWorkbook"), Global, "Excel.Global", "Workbook", false, false, Accessibility.Global, DeclarationType.PropertyGet);
        }
    }
}