using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.Concrete
{
    using ControlSize         = RibbonControlSize;
    using LanguageStrings     = IRibbonTextLanguageControl;

    public class RibbonGroup : RibbonCommon
    {
        internal RibbonGroup(string id, LanguageStrings strings, bool visible, bool enabled, ControlSize size)
            : base(id, strings, visible, enabled, size) {; }
    }
}
