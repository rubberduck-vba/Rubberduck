using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;

namespace Rubberduck.VBEditor.Extensions
{
    public static class WindowExtensions
    {
        //The hWnd property on most of the VBE windows is completely worthless (it just returns 0).  This will get its
        //*real* hWnd.
        public static IntPtr ToHwnd(this IWindow window)
        {
            //Try the obvious first.
            if (window.HWnd != 0)
            {
                return new IntPtr(window.HWnd);
            }
            //Try FindWindowEx next.
            var hwnd = NativeMethods.FindWindowEx(IntPtr.Zero, IntPtr.Zero, window.Type.ToClassName(), window.Caption);
            if (hwnd != IntPtr.Zero)
            {
                return hwnd;
            }
            //Try enum all the child windows last - ugh.
            var finder = new NativeMethods.ChildWindowFinder(window.Caption);
            finder.EnumWindowsProcToChildWindowByCaption(IntPtr.Zero, IntPtr.Zero);
            return finder.ResultHandle;
        }

        public static string ToClassName(this WindowKind kind)
        {
            switch (kind)
            {
                case WindowKind.Designer:
                    return "DesignerWindow";
                case WindowKind.Browser:
                    return "DockingView";
                case WindowKind.CodeWindow:
                case WindowKind.Watch:
                case WindowKind.Locals:
                case WindowKind.Immediate:
                    return "VbaWindow";
                case WindowKind.ProjectWindow:
                    return "PROJECT";
                case WindowKind.PropertyWindow:
                    return "wndclass_pbrs";
                case WindowKind.Find:                    
                case WindowKind.FindReplace:
                    return "#32770 (Dialog)";
                case WindowKind.Toolbox:
                    return "F3 MinFrame 77ec0000";
                case WindowKind.LinkedWindowFrame:  //WTF is this?
                    return string.Empty;        
                case WindowKind.MainWindow:
                    return "wndclass_desked_gsk";
                case WindowKind.ToolWindow:
                    return "VBFloatingPalette";
                default:
                    return string.Empty;       
            }
        }
    }
}
