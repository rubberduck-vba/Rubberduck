////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////


namespace Rubberduck.RibbonDispatcher {
    internal static class DispIds {
        public const int RibbonFactory      =  1;   // No conflict with "Id" below as only in disjoint interfaces.
        public const int ControlId          =  1;   // No conflict with "Id" below as only in disjoint interfaces.
        public const int Id                 =  1;

        // IRibbonCommon
        public const int Description        =  2;
        public const int KeyTip             =  3;
        public const int Label              =  4;
        public const int ScreenTip          =  5;
        public const int SuperTip           =  6;
        public const int SetLanguageStrings =  7;
        public const int IsEnabled          =  8;
        public const int IsVisible          =  9;
        public const int Size               = 10;

        // IImageableItem
        public const int Image              = 11;
        public const int ShowImage          = 12;
        public const int ShowLabel          = 13;
        public const int SetImage           = 14;
        public const int SetImageMso        = 15;

        // IActionItem
        public const int OnAction           = 16;

        // IToggleItem
        public const int IsPressed          = 17;
        public const int OnActionToggle     = 18;

        // IDropDownItem
        public const int SelectedItemId     = 19;
        public const int OnActionDropDown   = 20;
        //public const int Image              = 21;
        //public const int ShowImage          = 22;
        //public const int ShowLabel          = 23;
        //public const int IsVisible          = 24;
        //public const int Pressed            = 25;
        //public const int OnAction           = 26;
        //public const int OnActionToggle     = 27;
        //public const int SetImage           = 28;
        //public const int SetImageMso        = 29;
    }
}
