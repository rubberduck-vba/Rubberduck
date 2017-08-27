using System.Collections.Generic;

// I have left these three collected together for now, until I complete the design and implementation
// of their concrete classes - as a convenience and reminder.
namespace Rubberduck.RibbonDispatcher.Abstract
{
    public interface IRibbonText : IReadOnlyList<IRibbonTextLanguage>
    {
        string DefaultLangCode { get; }
    }

    /// <summary>All the control-specific {IRibbonTextLanguageControl} for a language.</summary>
    public interface IRibbonTextLanguage : IReadOnlyList<IRibbonTextLanguageControl>
    {
        string LangCode         { get; }
    }

    /// <summary>All the language-specific {IRibbonTextLanguageControl} for a ribbon control.</summary>
    public interface IRibbonTextControl : IReadOnlyList<IRibbonTextLanguageControl>
    {
        string ControlId        { get; }
    }
}
