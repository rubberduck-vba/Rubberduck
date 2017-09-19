namespace Rubberduck.RibbonDispatcher {
    internal static class DispIds {
        public const int RibbonFactory          =  1;   // No conflict with "Id" below as only in disjoint interfaces.
        public const int ControlId              =  1;   // No conflict with "Id" below as only in disjoint interfaces.
        public const int Id                     =  1;

        // IRibbonCommon
        public const int Description            =  2;
        public const int KeyTip                 =  3;
        public const int Label                  =  4;
        public const int ScreenTip              =  5;
        public const int SuperTip               =  6;
        public const int SetLanguageStrings     =  7;
        public const int IsEnabled              =  8;
        public const int IsVisible              =  9;
        public const int Size                   = 10;

        // IImageableItem
        public const int Image                  = 11;
        public const int ShowImage              = 12;
        public const int ShowLabel              = 13;
        public const int SetImageDisp           = 14;
        public const int SetImageMso            = 15;

        // IActionItem
        public const int OnAction               = 16;

        // IToggleItem
        public const int IsPressed              = 17;
        public const int OnActionToggle         = 18;

        // IDropDownItem
        public const int SelectedItemId         = 19;
        public const int SelectedItemIndex      = 20;
        public const int OnActionDropDown       = 21;

        public const int AddItem                = 22;
        public const int AddItemMso             = 23;

        public const int ItemCount              = 24;
        public const int ItemId                 = 25;
        public const int ItemLabel              = 26;
        public const int ItemScreenTip          = 27;
        public const int ItemSuperTip           = 28;
        public const int ItemImage              = 29;
        public const int ItemShowLabel          = 30;
        public const int ItemShowImage          = 31;

        public const int LoadImage              = 32;

        // COM Events
        public const int OnClicked              = 33;
        public const int OnToggled              = 34;
        public const int OnSelected             = 35;

        // Ribbon COmmands
        public const int Invalidate             = 41;
        public const int InvalidateControl      = 42;
        public const int InvalidateControlMso   = 43;
        public const int ActivateTab            = 44;
    }
}
