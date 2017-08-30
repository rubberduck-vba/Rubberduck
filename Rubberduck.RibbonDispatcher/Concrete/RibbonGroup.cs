using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.Concrete
{
    using LanguageStrings     = IRibbonTextLanguageControl;

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class RibbonGroup : RibbonCommon
    {
        internal RibbonGroup(string id, LanguageStrings strings, bool visible, bool enabled, MyRibbonControlSize size)
            : base(id, strings, visible, enabled, size) {; }
    }
}
