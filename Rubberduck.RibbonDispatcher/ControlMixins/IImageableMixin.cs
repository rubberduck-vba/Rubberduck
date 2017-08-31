////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using stdole;

namespace Rubberduck.RibbonDispatcher.ControlMixins {
    /// <summary>The interface for controls that can display an Image.</summary>
    [CLSCompliant(true)]
    public interface IImageableMixin {
        /// <summary>TODO</summary>
        [DispId(DispIds.Image)]
        object  Image     { get; }
        /// <summary>Returns or set whether to show the control's image; ignored by Large controls.</summary>
        [DispId(DispIds.ShowImage)]
        bool    ShowImage { get; set; }
        /// <summary>Returns or set whether to show the control's label; ignored by Large controls.</summary>
        [DispId(DispIds.ShowLabel)]
        bool    ShowLabel { get; set; }

        /// <summary>TODO</summary>
        [DispId(DispIds.SetImage)]
        void    SetImage(IPictureDisp image);
        /// <summary>TODO</summary>
        [DispId(DispIds.SetImageMso)]
        void    SetImageMso(string imageMso);
    }
}
