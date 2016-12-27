using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace Rubberduck.VBEditor.Native
{
    public static class WinEvents
    {
        #region Debugging symbol lookups
        public static readonly Dictionary<uint, string> EventNameLookup = new Dictionary<uint, string>
        {
            {0x1, "EVENT_SYSTEM_SOUND"},
            {0x2, "EVENT_SYSTEM_ALERT"},
            {0x3, "EVENT_SYSTEM_FOREGROUND"},
            {0x4, "EVENT_SYSTEM_MENUSTART"},
            {0x5, "EVENT_SYSTEM_MENUEND"},
            {0x6, "EVENT_SYSTEM_MENUPOPUPSTART"},
            {0x7, "EVENT_SYSTEM_MENUPOPUPEND"},
            {0x8, "EVENT_SYSTEM_CAPTURESTART"},
            {0x9, "EVENT_SYSTEM_CAPTUREEND"},
            {0xa, "EVENT_SYSTEM_MOVESIZESTART"},
            {0xb, "EVENT_SYSTEM_MOVESIZEEND"},
            {0xc, "EVENT_SYSTEM_CONTEXTHELPSTART"},
            {0xd, "EVENT_SYSTEM_CONTEXTHELPEND"},
            {0xe, "EVENT_SYSTEM_DRAGDROPSTART"},
            {0xf, "EVENT_SYSTEM_DRAGDROPEND"},
            {0x10, "EVENT_SYSTEM_DIALOGSTART"},
            {0x11, "EVENT_SYSTEM_DIALOGEND"},
            {0x12, "EVENT_SYSTEM_SCROLLINGSTART"},
            {0x13, "EVENT_SYSTEM_SCROLLINGEND"},
            {0x14, "EVENT_SYSTEM_SWITCHSTART"},
            {0x15, "EVENT_SYSTEM_SWITCHEND"},
            {0x16, "EVENT_SYSTEM_MINIMIZESTART"},
            {0x17, "EVENT_SYSTEM_MINIMIZEEND"},
            {0x8000, "EVENT_OBJECT_CREATE"},
            {0x8001, "EVENT_OBJECT_DESTROY"},
            {0x8002, "EVENT_OBJECT_SHOW"},
            {0x8003, "EVENT_OBJECT_HIDE"},
            {0x8004, "EVENT_OBJECT_REORDER"},
            {0x8005, "EVENT_OBJECT_FOCUS"},
            {0x8006, "EVENT_OBJECT_SELECTION"},
            {0x8007, "EVENT_OBJECT_SELECTIONADD"},
            {0x8008, "EVENT_OBJECT_SELECTIONREMOVE"},
            {0x8009, "EVENT_OBJECT_SELECTIONWITHIN"},
            {0x800A, "EVENT_OBJECT_STATECHANGE"},
            {0x800B, "EVENT_OBJECT_LOCATIONCHANGE"},
            {0x800C, "EVENT_OBJECT_NAMECHANGE"},
            {0x800D, "EVENT_OBJECT_DESCRIPTIONCHANGE"},
            {0x800E, "EVENT_OBJECT_VALUECHANGE"},
            {0x800F, "EVENT_OBJECT_PARENTCHANGE"},
            {0x8010, "EVENT_OBJECT_HELPCHANGE"},
            {0x8011, "EVENT_OBJECT_DEFACTIONCHANGE"},
            {0x8012, "EVENT_OBJECT_ACCELERATORCHANGE"},
        };

        public static readonly Dictionary<uint, string> ObjectIdNameLookup = new Dictionary<uint, string>
        {
            { 0x00000000, "OBJID_WINDOW" },
            { 0xFFFFFFFF, "OBJID_SYSMENU" },
            { 0xFFFFFFFE, "OBJID_TITLEBAR" },
            { 0xFFFFFFFD, "OBJID_MENU" },
            { 0xFFFFFFFC, "OBJID_CLIENT" },
            { 0xFFFFFFFB, "OBJID_VSCROLL" },
            { 0xFFFFFFFA, "OBJID_HSCROLL" },
            { 0xFFFFFFF9, "OBJID_SIZEGRIP" },
            { 0xFFFFFFF8, "OBJID_CARET" },
            { 0xFFFFFFF7, "OBJID_CURSOR" },
            { 0xFFFFFFF6, "OBJID_ALERT" },
            { 0xFFFFFFF5, "OBJID_SOUND" },
            { 0xFFFFFFF4, "OBJID_QUERYCLASSNAMEIDX" },
            { 0xFFFFFFF0, "OBJID_NATIVEOM" },        
        };
        #endregion

        #region System enumerations
        // ReSharper disable InconsistentNaming
        //See https://msdn.microsoft.com/en-us/library/windows/desktop/dd318066(v=vs.85).aspx
        public enum EventConstant : uint
        {
            EVENT_MIN = 0x1,
            EVENT_SYSTEM_SOUND = 0x1,
            EVENT_SYSTEM_ALERT = 0x2,
            EVENT_SYSTEM_FOREGROUND = 0x3,
            EVENT_SYSTEM_MENUSTART = 0x4,
            EVENT_SYSTEM_MENUEND = 0x5,
            EVENT_SYSTEM_MENUPOPUPSTART = 0x6,
            EVENT_SYSTEM_MENUPOPUPEND = 0x7,
            EVENT_SYSTEM_CAPTURESTART = 0x8,
            EVENT_SYSTEM_CAPTUREEND = 0x9,
            EVENT_SYSTEM_MOVESIZESTART = 0xa,
            EVENT_SYSTEM_MOVESIZEEND = 0xb,
            EVENT_SYSTEM_CONTEXTHELPSTART = 0xc,
            EVENT_SYSTEM_CONTEXTHELPEND = 0xd,
            EVENT_SYSTEM_DRAGDROPSTART = 0xe,
            EVENT_SYSTEM_DRAGDROPEND = 0xf,
            EVENT_SYSTEM_DIALOGSTART = 0x10,
            EVENT_SYSTEM_DIALOGEND = 0x11,
            EVENT_SYSTEM_SCROLLINGSTART = 0x12,
            EVENT_SYSTEM_SCROLLINGEND = 0x13,
            EVENT_SYSTEM_SWITCHSTART = 0x14,
            EVENT_SYSTEM_SWITCHEND = 0x15,
            EVENT_SYSTEM_MINIMIZESTART = 0x16,
            EVENT_SYSTEM_MINIMIZEEND = 0x17,
            EVENT_OEM_DEFINED_START = 0x0101,
            EVENT_OEM_DEFINED_END = 0x01FF,
            EVENT_AIA_START = 0xA000,
            EVENT_AIA_END = 0xAFFF,
            EVENT_UIA_EVENTID_START = 0x4E00,
            EVENT_UIA_EVENTID_END = 0x4EFF,
            EVENT_UIA_PROPID_START = 0x7500,
            EVENT_UIA_PROPID_END = 0x75FF,
            EVENT_OBJECT_START = 0x8000,
            EVENT_OBJECT_CREATE = 0x8000,
            EVENT_OBJECT_DESTROY = 0x8001,
            EVENT_OBJECT_SHOW = 0x8002,
            EVENT_OBJECT_HIDE = 0x8003,
            EVENT_OBJECT_REORDER = 0x8004,
            EVENT_OBJECT_FOCUS = 0x8005,
            EVENT_OBJECT_SELECTION = 0x8006,
            EVENT_OBJECT_SELECTIONADD = 0x8007,
            EVENT_OBJECT_SELECTIONREMOVE = 0x8008,
            EVENT_OBJECT_SELECTIONWITHIN = 0x8009,
            EVENT_OBJECT_STATECHANGE = 0x800A,
            EVENT_OBJECT_LOCATIONCHANGE = 0x800B,
            EVENT_OBJECT_NAMECHANGE = 0x800C,
            EVENT_OBJECT_DESCRIPTIONCHANGE = 0x800D,
            EVENT_OBJECT_VALUECHANGE = 0x800E,
            EVENT_OBJECT_PARENTCHANGE = 0x800F,
            EVENT_OBJECT_HELPCHANGE = 0x8010,
            EVENT_OBJECT_DEFACTIONCHANGE = 0x8011,
            EVENT_OBJECT_ACCELERATORCHANGE = 0x8012,
            EVENT_OBJECT_INVOKED = 0x8013,
            EVENT_OBJECT_CONTENTSCROLLED = 0x8015,
            EVENT_SYSTEM_ARRANGMENTPREVIEW = 0x8016,
            EVENT_OBJECT_LIVEREGIONCHANGED = 0x8019,
            EVENT_OBJECT_HOSTEDOBJECTSINVALIDATED = 0x8020,
            EVENT_OBJECT_DRAGSTART = 0x8021,
            EVENT_OBJECT_DRAGCANCEL = 0x8022,
            EVENT_OBJECT_DRAGCOMPLETE = 0x8023,
            EVENT_OBJECT_DRAGENTER = 0x8024,
            EVENT_OBJECT_DRAGLEAVE =  0x8025,
            EVENT_OBJECT_DRAGDROPPED = 0x8026,
            EVENT_OBJECT_IME_SHOW = 0x8027,
            EVENT_OBJECT_IME_HIDE = 0x8028,
            EVENT_OBJECT_IME_CHANGE = 0x8029,
            EVENT_OBJECT_TEXTEDIT_CONVERSIONTARGETCHANGED = 0x8030,
            EVENT_OBJECT_TEXTSELECTIONCHANGED = 0x8014,
            EVENT_OBJECT_END = 0x80FF,
            EVENT_MAX = 0x7FFFFFFF
        }
        // possible marshaling unmanaged type conflict/problem between 32/64 bit

        public enum ObjId : uint
        {
            OBJID_WINDOW = 0x00000000,
            OBJID_SYSMENU = 0xFFFFFFFF,
            OBJID_TITLEBAR = 0xFFFFFFFE,
            OBJID_MENU = 0xFFFFFFFD,
            OBJID_CLIENT = 0xFFFFFFFC,
            OBJID_VSCROLL = 0xFFFFFFFB,
            OBJID_HSCROLL = 0xFFFFFFFA,
            OBJID_SIZEGRIP = 0xFFFFFFF9,
            OBJID_CARET = 0xFFFFFFF8,
            OBJID_CURSOR = 0xFFFFFFF7,
            OBJID_ALERT = 0xFFFFFFF6,
            OBJID_SOUND = 0xFFFFFFF5,
            OBJID_QUERYCLASSNAMEIDX = 0xFFFFFFF4,
            OBJID_NATIVEOM = 0xFFFFFFF0
        }

        public enum WinEventFlags : uint
        {
            WINEVENT_OUTOFCONTEXT = 0x0000,
            WINEVENT_SKIPOWNTHREAD = 0x0001,
            WINEVENT_SKIPOWNPROCESS = 0x0002,
            WINEVENT_INCONTEXT = 0x0004
        }
        // ReSharper restore InconsistentNaming
        #endregion

        #region API declarations

        public delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, uint idObject, uint idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        public static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, IntPtr lpfnWinEventProc, uint idProcess,
                                             uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        public static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        #endregion

        #region Extension methods

        public static string ToObjectIdString(this uint objectId)
        {
            return ObjectIdNameLookup.ContainsKey(objectId) ? ObjectIdNameLookup[objectId] : objectId.ToString(CultureInfo.InvariantCulture);
        }

        public static string ToEventIdString(this uint eventId)
        {
            return EventNameLookup.ContainsKey(eventId) ? EventNameLookup[eventId] : eventId.ToString(CultureInfo.InvariantCulture);
        }

        public static string ToClassName(this IntPtr hwnd)
        {
            var buffer = new StringBuilder(256);
            if (hwnd != IntPtr.Zero)
            {
                return GetClassName(hwnd, buffer, buffer.Capacity) != 0 ? buffer.ToString() : string.Empty;
            }
            return string.Empty;
        }

        #endregion
    }
}
