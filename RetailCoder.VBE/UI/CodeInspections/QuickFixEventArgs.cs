using System;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.CodeInspections
{
    [ComVisible(false)]
    public class QuickFixEventArgs
    {
        private readonly Action<VBE> _quickFix;

        public QuickFixEventArgs(Action<VBE> quickFix)
        {
            _quickFix = quickFix;
        }

        public Action<VBE> QuickFix
        {
            get { return _quickFix; }
        }
    }
}