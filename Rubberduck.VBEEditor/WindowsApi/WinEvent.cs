using System.Collections.Generic;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.VBEditor.WindowsApi
{
    public enum WinEvent
    {
        [ReflectionIgnore]
        Min = 0x0001,
        SystemSound = 0x0001,
        SystemAlert = 0x0002,
        SystemForeground = 0x0003,
        SystemMenuStart = 0x0004,
        SystemMenuEnd = 0x0005,
        SystemMenuPopupStart = 0x0006,
        SystemMenuPopupEnd = 0x0007,
        SystemCaptureStart = 0x0008,
        SystemCaptureEnd = 0x0009,
        SystemMoveSizeStart = 0x000A,
        SystemMoveSizeEnd = 0x000B,
        SystemContextHelpStart = 0x000C,
        SystemContextHelpEnd = 0x000D,
        SystemDragDropStart = 0x000E,
        SystemDragDropEnd = 0x000F,
        SystemDialogStart = 0x0010,
        SystemDialogEnd = 0x0011,
        SystemScrollingStart = 0x0012,
        SystemScrollingEnd = 0x0013,
        SystemSwitchStart = 0x0014,
        SystemSwitchEnd = 0x0015,
        SystemMinimizeStart = 0x0016,
        SystemMinimizeEnd = 0x0017,
        SystemDesktopSwitch = 0x0020,
        SystemEnd = 0x00FF,
        ObjectCreate = 0x8000,
        ObjectDestroy = 0x8001,
        ObjectShow = 0x8002,
        ObjectHide = 0x8003,
        ObjectReorder = 0x8004,
        ObjectFocus = 0x8005,
        ObjectSelection = 0x8006,
        ObjectSelectionAdd = 0x8007,
        ObjectSelectionRemove = 0x8008,
        ObjectSelectionWithin = 0x8009,
        ObjectStateChange = 0x800A,
        ObjectLocationChange = 0x800B,
        ObjectNameChange = 0x800C,
        ObjectDescriptionChange = 0x800D,
        ObjectValueChange = 0x800E,
        ObjectParentChange = 0x800F,
        ObjectHelpChange = 0x8010,
        ObjectDefactionChange = 0x8011,
        ObjectInvoked = 0x8013,
        ObjectTextSelectionChanged = 0x8014,
        SystemArrangmentPreview = 0x8016,
        ObjectLiveRegionChanged = 0x8019,
        ObjectHostedObjectsInvalidated = 0x8020,
        ObjectDragStart = 0x8021,
        ObjectDragCancel = 0x8022,
        ObjectDragComplete = 0x8023,
        ObjectDragEnter = 0x8024,
        ObjectDragLeave = 0x8025,
        ObjectDragDropped = 0x8026,
        ObjectImeShow = 0x8027,
        ObjectImeHide = 0x8028,
        ObjectImeChange = 0x8029,
        ObjectTexteditConversionTargetChanged = 0x8030,
        ObjectEnd = 0x80FF,
        AiaStart = 0xA000,
        AiaEnd = 0xAFFF,
        [ReflectionIgnore]
        Max = 0x7FFFFFFF
    }

    public enum ObjId
    {
        IntelliSense = 3,
        Window = 0,
        SysMenu = -1,
        TitleBar = -2,
        Menu = -3,
        Client = -4,
        VScroll = -5,
        HScroll = -6,
        SizeGrip = -7,
        Caret = -8,
        Cursor = -9,
        Alert = -10,
        Sount = -11,
        QueryClasNameIdx = -12,
        NativeOM = -16     
    }

    public static class WinEventMap
    {
        public static readonly Dictionary<int, string> Lookup = EnumHelper.ToDictionary<WinEvent, int>();
    }
}
