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
        public static readonly QualifiedModuleName WorkbookModuleName = new QualifiedModuleName("Excel", "Workbook");
        public static readonly QualifiedModuleName WorksheetModuleName = new QualifiedModuleName("Excel", "Worksheet");

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
            public static Declaration Application = new Declaration(new QualifiedMemberName(ExcelModuleName, "Application"), Excel, "Excel.Application", "Application", false, false, Accessibility.Global, DeclarationType.PropertyGet);

        }

        private class WorkbookClass
        {
            public static readonly Declaration Workbook = new Declaration(new QualifiedMemberName(ExcelModuleName, "Workbook"), ExcelLib.Excel, "Excel", "Workbook", false, false, Accessibility.Global, DeclarationType.Class);

            public static Declaration ActiveSheet = new Declaration(new QualifiedMemberName(WorkbookModuleName, "ActiveSheet"), Workbook, "Excel.Workbook", "Worksheet", false, false, Accessibility.Public, DeclarationType.PropertyGet); // cheating on return type
            public static Declaration Sheets = new Declaration(new QualifiedMemberName(WorkbookModuleName, "Sheets"), Workbook, "Excel.Workbook", "Sheets", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration Worksheets = new Declaration(new QualifiedMemberName(WorkbookModuleName, "Worksheets"), Workbook, "Excel.Workbook", "Worksheets", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration Names = new Declaration(new QualifiedMemberName(WorkbookModuleName, "Names"), Workbook, "Excel.Workbook", "Names", false, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static Declaration Activate = new Declaration(new QualifiedMemberName(WorkbookModuleName, "Activate"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration AddinInstall = new Declaration(new QualifiedMemberName(WorkbookModuleName, "AddinInstall"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration AddinUninstall = new Declaration(new QualifiedMemberName(WorkbookModuleName, "AddinUninstall"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration AfterSave = new Declaration(new QualifiedMemberName(WorkbookModuleName, "AfterSave"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration AfterXmlExport = new Declaration(new QualifiedMemberName(WorkbookModuleName, "AfterXmlExport"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration AfterXmlImport = new Declaration(new QualifiedMemberName(WorkbookModuleName, "AfterXmlImport"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration BeforeClose = new Declaration(new QualifiedMemberName(WorkbookModuleName, "BeforeClose"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration BeforePrint = new Declaration(new QualifiedMemberName(WorkbookModuleName, "BeforePrint"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration BeforeSave = new Declaration(new QualifiedMemberName(WorkbookModuleName, "BeforeSave"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration BeforeXmlExport = new Declaration(new QualifiedMemberName(WorkbookModuleName, "BeforeXmlExport"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration BeforeXmlImport = new Declaration(new QualifiedMemberName(WorkbookModuleName, "BeforeXmlImport"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Deactivate = new Declaration(new QualifiedMemberName(WorkbookModuleName, "Deactivate"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration NewChart = new Declaration(new QualifiedMemberName(WorkbookModuleName, "NewChart"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration NewSheet = new Declaration(new QualifiedMemberName(WorkbookModuleName, "NewSheet"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Open = new Declaration(new QualifiedMemberName(WorkbookModuleName, "Open"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration PivotTableCloseConnection = new Declaration(new QualifiedMemberName(WorkbookModuleName, "PivotTableCloseConnection"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration PivotTableOpenConnection = new Declaration(new QualifiedMemberName(WorkbookModuleName, "PivotTableOpenConnection"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration RowsetComplete = new Declaration(new QualifiedMemberName(WorkbookModuleName, "RowsetComplete"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetActivate = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetActivate"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetBeforeDoubleClick = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetBeforeDoubleClick"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetBeforeRightClick = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetBeforeRightClick"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetCalculate = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetCalculate"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetChange = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetChange"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetDeactivate = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetDeactivate"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetFollowHyperlink = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetFollowHyperlink"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetPivotTableAfterValueChange = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetPivotTableAfterValueChange"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetPivotTableBeforeAllocateChanges = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetPivotTableBeforeAllocateChanges"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetPivotTableBeforeCommitChanges = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetPivotTableBeforeCommitChanges"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetPivotTableBeforeDiscardChanges = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetPivotTableBeforeDiscardChanges"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetPivotTableChangeSync = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetPivotTableChangeSync"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetPivotTableUpdate = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetPivotTableUpdate"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SheetSelectionChange = new Declaration(new QualifiedMemberName(WorkbookModuleName, "SheetSelectionChange"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Sync = new Declaration(new QualifiedMemberName(WorkbookModuleName, "Sync"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration WindowActivate = new Declaration(new QualifiedMemberName(WorkbookModuleName, "WindowActivate"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration WindowDeactivate = new Declaration(new QualifiedMemberName(WorkbookModuleName, "WindowDeactivate"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration WindowResize = new Declaration(new QualifiedMemberName(WorkbookModuleName, "WindowResize"), Workbook, "Excel.Workbook", null, false, false, Accessibility.Public, DeclarationType.Event);
        }

        private class WorksheetClass
        {
            public static readonly Declaration Worksheet = new Declaration(new QualifiedMemberName(ExcelModuleName, "Worksheet"), ExcelLib.Excel, "Excel", "Worksheet", false, false, Accessibility.Global, DeclarationType.Class);

            public static Declaration Evaluate = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Evaluate"), Worksheet, "Excel.Worksheet", "Variant", false, false, Accessibility.Public, DeclarationType.Function);
            public static Declaration Range = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Range"), Worksheet, "Excel.Worksheet", "Range", true, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration RangeAssign = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Range"), Worksheet, "Excel.Worksheet", "Range", true, false, Accessibility.Public, DeclarationType.PropertyLet); // cheating

            public static Declaration Activate = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Activate"), Worksheet, "Excel.Worksheet", "Variant", false, false, Accessibility.Public, DeclarationType.Function);
            public static Declaration Cells = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Cells"), Worksheet, "Excel.Worksheet", "Range", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration CellsAssign = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Cells"), Worksheet, "Excel.Worksheet", "Range", false, false, Accessibility.Public, DeclarationType.PropertyLet); // cheating to simulate default property of return type.
            public static Declaration Names = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Names"), Worksheet, "Excel.Worksheet", "Names", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration WorksheetFunction = new Declaration(new QualifiedMemberName(WorksheetModuleName, "WorksheetFunction"), Worksheet, "Excel.Worksheet", "WorksheetFunction", false, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static Declaration Columns = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Columns"), Worksheet, "Excel.Worksheet", "Range", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration Rows = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Rows"), Worksheet, "Excel.Worksheet", "Range", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration ActiveCell = new Declaration(new QualifiedMemberName(WorksheetModuleName, "ActiveCell"), Worksheet, "Excel.Worksheet", "Range", false, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static Declaration ActivateEvent = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Activate"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration BeforeDoubleClick = new Declaration(new QualifiedMemberName(WorksheetModuleName, "BeforeDoubleClick"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration BeforeRightClick = new Declaration(new QualifiedMemberName(WorksheetModuleName, "BeforeRightClick"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Calculate = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Calculate"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Change = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Change"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Deactivate = new Declaration(new QualifiedMemberName(WorksheetModuleName, "Deactivate"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration FollowHyperlink = new Declaration(new QualifiedMemberName(WorksheetModuleName, "FollowHyperlink"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration PivotTableAfterValueChange = new Declaration(new QualifiedMemberName(WorksheetModuleName, "PivotTableAfterValueChange"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration PivotTableBeforeAllocateChanges = new Declaration(new QualifiedMemberName(WorksheetModuleName, "PivotTableBeforeAllocateChanges"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration PivotTableBeforeCommitChanges = new Declaration(new QualifiedMemberName(WorksheetModuleName, "PivotTableBeforeCommitChanges"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration PivotTableBeforeDiscardChanges = new Declaration(new QualifiedMemberName(WorksheetModuleName, "PivotTableBeforeDiscardChanges"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration PivotTableChangeSync = new Declaration(new QualifiedMemberName(WorksheetModuleName, "PivotTableChangeSync"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration PivotTableUpdate = new Declaration(new QualifiedMemberName(WorksheetModuleName, "PivotTableUpdate"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration SelectionChange = new Declaration(new QualifiedMemberName(WorksheetModuleName, "SelectionChange"), Worksheet, "Excel.Worksheet", null, false, false, Accessibility.Public, DeclarationType.Event);
        }

        private class RangeClass
        {
            public static readonly Declaration Range = new Declaration(new QualifiedMemberName(ExcelModuleName, "Range"), ExcelLib.Excel, "Excel", "Range", false, false, Accessibility.Global, DeclarationType.Class);

            public static Declaration Cells = new Declaration(new QualifiedMemberName(ExcelModuleName, "Cells"), Range, "Excel.Range", "Range", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration CellsAssign = new Declaration(new QualifiedMemberName(ExcelModuleName, "Cells"), Range, "Excel.Range", "Range", false, false, Accessibility.Public, DeclarationType.PropertyLet); // cheating to simulate default property of return type.
            public static Declaration Activate = new Declaration(new QualifiedMemberName(ExcelModuleName, "Activate"), Range, "Excel.Range", "Variant", false, false, Accessibility.Public, DeclarationType.Function);
            public static Declaration Select = new Declaration(new QualifiedMemberName(ExcelModuleName, "Select"), Range, "Excel.Range", "Variant", false, false, Accessibility.Public, DeclarationType.Function);
            public static Declaration Columns = new Declaration(new QualifiedMemberName(ExcelModuleName, "Columns"), Range, "Excel.Range", "Range", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration Rows = new Declaration(new QualifiedMemberName(ExcelModuleName, "Rows"), Range, "Excel.Range", "Range", false, false, Accessibility.Public, DeclarationType.PropertyGet);
        }

        private class GlobalClass
        {
            public static readonly Declaration Global = new Declaration(new QualifiedMemberName(ExcelModuleName, "Global"), ExcelLib.Excel, "Excel", "Global", false, false, Accessibility.Global, DeclarationType.Module); // cheating, it's actually a class.

            public static Declaration Evaluate = new Declaration(new QualifiedMemberName(ExcelModuleName, "Evaluate"), Global, "Excel.Global", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Range = new Declaration(new QualifiedMemberName(ExcelModuleName, "Range"), Global, "Excel.Global", "Range", true, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration RangeAssign = new Declaration(new QualifiedMemberName(ExcelModuleName, "Range"), Global, "Excel.Global", "Range", true, false, Accessibility.Global, DeclarationType.PropertyLet); // cheating to simuate default property of return type.
            public static Declaration Selection = new Declaration(new QualifiedMemberName(ExcelModuleName, "Selection"), Global, "Excel.Global", "Object", true, false, Accessibility.Global, DeclarationType.PropertyGet);

            public static Declaration Activate = new Declaration(new QualifiedMemberName(ExcelModuleName, "Activate"), Global, "Excel.Global", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Select = new Declaration(new QualifiedMemberName(ExcelModuleName, "Select"), Global, "Excel.Global", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Cells = new Declaration(new QualifiedMemberName(ExcelModuleName, "Cells"), Global, "Excel.Global", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration CellsAssign = new Declaration(new QualifiedMemberName(ExcelModuleName, "Cells"), Global, "Excel.Global", "Range", false, false, Accessibility.Global, DeclarationType.PropertyLet); // cheating to simulate default property of return type.
            public static Declaration Names = new Declaration(new QualifiedMemberName(ExcelModuleName, "Names"), Global, "Excel.Global", "Names", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Sheets = new Declaration(new QualifiedMemberName(ExcelModuleName, "Sheets"), Global, "Excel.Global", "Sheets", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Worksheets = new Declaration(new QualifiedMemberName(ExcelModuleName, "Worksheets"), Global, "Excel.Global", "Worksheets", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration WorksheetFunction = new Declaration(new QualifiedMemberName(ExcelModuleName, "WorksheetFunction"), Global, "Excel.Global", "WorksheetFunction", false, false, Accessibility.Global, DeclarationType.PropertyGet);

            public static Declaration Columns = new Declaration(new QualifiedMemberName(ExcelModuleName, "Columns"), Global, "Excel.Global", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Rows = new Declaration(new QualifiedMemberName(ExcelModuleName, "Rows"), Global, "Excel.Global", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);

            public static Declaration ActiveCell = new Declaration(new QualifiedMemberName(ExcelModuleName, "ActiveCell"), Global, "Excel.Global", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration ActiveSheet = new Declaration(new QualifiedMemberName(ExcelModuleName, "ActiveSheet"), Global, "Excel.Global", "Worksheet", false, false, Accessibility.Global, DeclarationType.PropertyGet); // cheating on return type
            public static Declaration ActiveWorkbook = new Declaration(new QualifiedMemberName(ExcelModuleName, "ActiveWorkbook"), Global, "Excel.Global", "Workbook", false, false, Accessibility.Global, DeclarationType.PropertyGet);
        }

        private class ApplicationClass
        {
            public static readonly Declaration Application = new Declaration(new QualifiedMemberName(ExcelModuleName, "Application"), ExcelLib.Excel, "Application", "Application", false, false, Accessibility.Global, DeclarationType.Module); // cheating, it's actually a class.

            public static Declaration Evaluate = new Declaration(new QualifiedMemberName(ExcelModuleName, "Evaluate"), Application, "Excel.Application", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Range = new Declaration(new QualifiedMemberName(ExcelModuleName, "Range"), Application, "Excel.Application", "Range", true, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration RangeAssign = new Declaration(new QualifiedMemberName(ExcelModuleName, "Range"), Application, "Excel.Application", "Range", true, false, Accessibility.Global, DeclarationType.PropertyLet); // cheating to simuate default property of return type.
            public static Declaration Selection = new Declaration(new QualifiedMemberName(ExcelModuleName, "Selection"), Application, "Excel.Application", "Object", true, false, Accessibility.Global, DeclarationType.PropertyGet);

            public static Declaration Activate = new Declaration(new QualifiedMemberName(ExcelModuleName, "Activate"), Application, "Excel.Application", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Select = new Declaration(new QualifiedMemberName(ExcelModuleName, "Select"), Application, "Excel.Application", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Cells = new Declaration(new QualifiedMemberName(ExcelModuleName, "Cells"), Application, "Excel.Application", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration CellsAssign = new Declaration(new QualifiedMemberName(ExcelModuleName, "Cells"), Application, "Excel.Application", "Range", false, false, Accessibility.Global, DeclarationType.PropertyLet); // cheating to simulate default property of return type.
            public static Declaration Names = new Declaration(new QualifiedMemberName(ExcelModuleName, "Names"), Application, "Excel.Application", "Names", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Sheets = new Declaration(new QualifiedMemberName(ExcelModuleName, "Sheets"), Application, "Excel.Application", "Sheets", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Worksheets = new Declaration(new QualifiedMemberName(ExcelModuleName, "Worksheets"), Application, "Excel.Application", "Worksheets", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration WorksheetFunction = new Declaration(new QualifiedMemberName(ExcelModuleName, "WorksheetFunction"), Application, "Excel.Application", "WorksheetFunction", false, false, Accessibility.Global, DeclarationType.PropertyGet);

            public static Declaration Columns = new Declaration(new QualifiedMemberName(ExcelModuleName, "Columns"), Application, "Excel.Application", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Rows = new Declaration(new QualifiedMemberName(ExcelModuleName, "Rows"), Application, "Excel.Application", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);

            public static Declaration ActiveCell = new Declaration(new QualifiedMemberName(ExcelModuleName, "ActiveCell"), Application, "Excel.Application", "Range", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration ActiveSheet = new Declaration(new QualifiedMemberName(ExcelModuleName, "ActiveSheet"), Application, "Excel.Application", "Worksheet", false, false, Accessibility.Global, DeclarationType.PropertyGet); // cheating on return type
            public static Declaration ActiveWorkbook = new Declaration(new QualifiedMemberName(ExcelModuleName, "ActiveWorkbook"), Application, "Excel.Application", "Workbook", false, false, Accessibility.Global, DeclarationType.PropertyGet);
        }
    }
}