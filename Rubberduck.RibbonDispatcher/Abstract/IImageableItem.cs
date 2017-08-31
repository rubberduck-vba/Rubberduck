////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using stdole;

namespace Rubberduck.RibbonDispatcher.Abstract {
    //[ComVisible(true)]
    //[Guid("42D56042-3FE9-4F1F-AD49-3ED0EE6CC987")]
    //[CLSCompliant(true)]
    //[InterfaceType(ComInterfaceType.InterfaceIsDual)]
    /// <summary>TODO</summary>
    public interface IImageableItem {
        /// <summary>TODO</summary>
        object  Image     { get; }
        /// <summary>Returns or set whether to show the control's image; ignored by Large controls.</summary>
        bool    ShowImage { get; set; }
        /// <summary>Returns or set whether to show the control's label; ignored by Large controls.</summary>
        bool    ShowLabel { get; set; }

        /// <summary>TODO</summary>
        void SetImage(IPictureDisp image);
        /// <summary>TODO</summary>
        void SetImageMso(string imageMso);
    }
}
