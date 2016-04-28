using System;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane
{
    public class CodePaneWrapper : ICodePaneWrapper
    {
        private readonly CodePane _codePane;
        public CodePane CodePane { get { return _codePane; } }

        public CodePanes Collection { get { return _codePane.Collection; } }
        public VBE VBE { get { return _codePane.VBE; } }
        public Window Window { get { return _codePane.Window; } }
        public int TopLine { 
            get { return _codePane.TopLine; } 
            set { _codePane.TopLine = value; }
        }
        public int CountOfVisibleLines { get { return _codePane == null ? 0 : _codePane.CountOfVisibleLines; } }
        public CodeModule CodeModule { get { return _codePane == null ? null : _codePane.CodeModule; } }
        public vbext_CodePaneview CodePaneView { get { return _codePane == null ? vbext_CodePaneview.vbext_cv_FullModuleView : _codePane.CodePaneView; } }

        public CodePaneWrapper(CodePane codePane)
        {
            // bug: if there's no active code pane, we're creating (and using) an invalid object -> NullReferenceException
            _codePane = codePane;
        }

        public void GetSelection(out int startLine, out int startColumn, out int endLine, out int endColumn)
        {
            _codePane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
        }

        public void SetSelection(int startLine, int startColumn, int endLine, int endColumn)
        {
            _codePane.SetSelection(startLine, startColumn, endLine, endColumn);
        }

        public void Show()
        {
            _codePane.Show();
        }

        public Selection Selection
        {
            get
            {
                try
                {
                    return GetSelection();
                }
                catch (COMException)
                {
                    // Gotcha
                }
                return new Selection();
            }

            set
            {
                SetSelection(value);
            }
        }

        private Selection GetSelection()
        {
            int startLine;
            int endLine;
            int startColumn;
            int endColumn;

            if (_codePane == null)
            {
                return new Selection();
            }

            GetSelection(out startLine, out startColumn, out endLine, out endColumn);

            if (endLine > startLine && endColumn == 1)
            {
                endLine--;
                endColumn = CodeModule.get_Lines(endLine, 1).Length;
            }

            return new Selection(startLine, startColumn, endLine, endColumn);
        }

        private void SetSelection(Selection selection)
        {
            SetSelection(selection.StartLine, selection.StartColumn, selection.EndLine, selection.EndColumn);
            ForceFocus();
        }

        /// <summary>   A CodePane extension method that forces focus onto the CodePane. This patches a bug in VBE.Interop.</summary>
        public void ForceFocus()
        {
            _codePane.ForceFocus();
        }
    }
}