using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Path = System.IO.Path;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Excel = Microsoft.Office.Interop.Excel;
using MSForms = Microsoft.Vbe.Interop.Forms;

namespace Rubberduck.VBEditor.Application
{
    public class ExcelApp : HostApplicationBase<Excel.Application>
    {
        public const int MaxPossibleLengthOfProcedureName = 255;

        public ExcelApp() : base("Excel") { }
        public ExcelApp(IVBE vbe) : base(vbe, "Excel") { }

        public override void Run(dynamic declaration)
        {
            var call = GenerateMethodCall(declaration);
            Application.Run(call);
        }

        public override object Run(string name, params object[] args)
        {
            switch (args.Length)
            {
                case 0:
                    return Application.Run(name);
                case 1:
                    return Application.Run(name, args[0]);
                case 2:
                    return Application.Run(name, args[0], args[1]);
                case 3:
                    return Application.Run(name, args[0], args[1], args[2]);
                case 4:
                    return Application.Run(name, args[0], args[1], args[2], args[3]);
                case 5:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4]);
                case 6:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5]);
                case 7:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6]);
                case 8:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7]);
                case 9:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8]);
                case 10:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9]);
                case 11:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10]);
                case 12:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11]);
                case 13:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12]);
                case 14:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13]);
                case 15:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14]);
                case 16:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15]);
                case 17:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16]);
                case 18:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17]);
                case 19:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18]);
                case 20:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19]);
                case 21:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19],args[20]);
                case 22:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19], args[20], args[21]);
                case 23:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19], args[20], args[21], args[22]);
                case 24:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19], args[20], args[21], args[22], args[23]);
                case 25:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19], args[20], args[21], args[22], args[23], args[24]);
                case 26:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19], args[20], args[21], args[22], args[23], args[24], args[25]);
                case 27:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19], args[20], args[21], args[22], args[23], args[24], args[25], args[26]);
                case 28:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19], args[20], args[21], args[22], args[23], args[24], args[25], args[26], args[27]);
                case 29:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19], args[20], args[21], args[22], args[23], args[24], args[25], args[26], args[27], args[28]);
                case 30:
                    return Application.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19], args[20], args[21], args[22], args[23], args[24], args[25], args[26], args[27], args[28], args[29]);
                default:
                    throw new ArgumentException("Too many arguments.");
            }
        }

        public override IEnumerable GetDocumentMacros()
        {
            var books = Application.Workbooks;
            foreach (Excel.Workbook workbook in books)
            {
                var worksheets = workbook.Worksheets;
                foreach (Excel.Worksheet worksheet in worksheets)
                {
                    foreach (var macro in GetMacrosFromShapes(worksheet))
                    {
                        yield return macro;
                    }

                    foreach (var macro in GetMacrosFromOleObjects(worksheet))
                    {
                        yield return macro;
                    }
                }

                Marshal.ReleaseComObject(worksheets);
            }
            Marshal.ReleaseComObject(books);
        }

        private static readonly string[] CommonEvents = new[]
        {
            "BeforeDragOver",
            "BeforeDropOrPaste",
            "Click",
            "DblClick",
            "Error",
            "KeyDown",
            "KeyPress",
            "KeyUp",
            "MouseDown",
            "MouseMove",
            "MouseUp",
        };

        private static IEnumerable<HostDocumentMacro> GetMacrosFromOleObjects(Excel.Worksheet worksheet)
        {
            var oleObjects = worksheet.OLEObjects();
            foreach (Excel.OLEObject oleObject in oleObjects)
            {
                yield return new HostDocumentMacro(oleObject.ShapeRange.ID, oleObject.Name, oleObject.progID, CommonEvents);

                var optionButton = oleObject.Object as MSForms.OptionButton;
                var checkBox = oleObject.Object as MSForms.CheckBox;
                var listBox = oleObject.Object as MSForms.ListBox;
                var toggleButton = oleObject.Object as MSForms.ToggleButton;

                var changingControls = new object[] 
                {
                    optionButton,
                    checkBox,
                    listBox,
                    toggleButton,
                };

                if (changingControls.Any(e => e != null))
                {
                    var events = CommonEvents.Union(new[] {nameof(optionButton.Change)});
                    yield return new HostDocumentMacro(oleObject.ShapeRange.ID, oleObject.Name, oleObject.progID, events);
                    continue;
                }

                var textBox = oleObject.Object as MSForms.TextBox;
                var comboBox = oleObject.Object as MSForms.ComboBoxClass;
                if ((textBox ?? (object)comboBox) != null)
                {
                    var events = CommonEvents.Union(new[] { "Change", "DropButtonClick" });
                    yield return new HostDocumentMacro(oleObject.ShapeRange.ID, oleObject.Name, oleObject.progID, events);
                    continue;
                }

                var spinButton = oleObject.Object as MSForms.SpinButton;
                var scrollBar = oleObject.Object as MSForms.ScrollBar;
                if ((spinButton ?? (object)scrollBar) != null)
                {
                    var events = CommonEvents
                        .Where(e => !e.StartsWith("Mouse") && !e.EndsWith("Click"))
                        .Union(new[]
                        {
                            nameof(spinButton.Change),
                            nameof(spinButton.SpinDown),
                            nameof(spinButton.SpinUp)
                        });
                    yield return new HostDocumentMacro(oleObject.ShapeRange.ID, oleObject.Name, oleObject.progID, events);
                    continue;
                }

                var multiPage = oleObject.Object as MSForms.MultiPage;
                var frame = oleObject.Object as MSForms.Frame;
                if ((multiPage ?? (object)frame) != null)
                {
                    var events = CommonEvents
                        .Where(e => !e.StartsWith("Mouse"))
                        .Union(new[]
                        {
                            nameof(frame.AddControl),
                            nameof(frame.Layout),
                            nameof(frame.RemoveControl),
                            nameof(frame.Scroll),
                            "Zoom" // nameof is ambiguous in this scope
                        });
                    yield return new HostDocumentMacro(oleObject.ShapeRange.ID, oleObject.Name, oleObject.progID, events);
                    continue;
                }
            }

            Marshal.ReleaseComObject(oleObjects);
        }

        private static IEnumerable GetMacrosFromShapes(Excel.Worksheet worksheet)
        {
            var shapes = worksheet.Shapes;
            foreach (Excel.Shape shape in shapes)
            {
                var actionName = shape.OnAction;
                if (!string.IsNullOrEmpty(actionName))
                {
                    if (shape.Type != MsoShapeType.msoOLEControlObject)
                    {
                        yield return new HostDocumentMacro(shape.ID, shape.Name, shape.OnAction);
                    }
                }
                Marshal.ReleaseComObject(shape);
            }
            Marshal.ReleaseComObject(shapes);
        }

        protected virtual string GenerateMethodCall(dynamic declaration)
        {
            var qualifiedMemberName = declaration.QualifiedName;
            var module = qualifiedMemberName.QualifiedModuleName;

            var documentName = string.IsNullOrEmpty(module.ProjectPath)
                ? declaration.ProjectDisplayName
                : Path.GetFileName(module.ProjectPath);

            var candidateString = string.IsNullOrEmpty(documentName)
                ? qualifiedMemberName.ToString()
                : string.Format("'{0}'!{1}", documentName.Replace("'", "''"), qualifiedMemberName);

            return candidateString.Length <= MaxPossibleLengthOfProcedureName
                ? candidateString
                : qualifiedMemberName.ToString();
        }
    }
}
