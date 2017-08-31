using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher.Concrete {

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IControlStrings))]
    [Guid(RubberduckGuid.ControlStrings)]
    public class ControlStrings : IControlStrings {
        /// <summary>TODO</summary>
        public ControlStrings() => _list = new Dictionary<string,string>();

        Dictionary<string,string> _list;

        /// <summary>TODO</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [SuppressMessage("Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.String.Format(System.String,System.Object)")]
        public IControlStrings AddControl(string ControlId,
            string Label          = null,
            string ScreenTip      = null,
            string SuperTip       = null,
            string AlternateLabel = null,
            string Description    = null,
            string KeyTip         = null) {
            _list.AddNotNull($"{ControlId}_Label",          Label);
            _list.AddNotNull($"{ControlId}_ScreenTip",      ScreenTip);
            _list.AddNotNull($"{ControlId}_SuperTip",       SuperTip);
            _list.AddNotNull($"{ControlId}_AlternateLabel", AlternateLabel);
            _list.AddNotNull($"{ControlId}_Description",    Description);
            _list.AddNotNull($"{ControlId}_KeyTip",         KeyTip);
             return this;
        }

        /// <inheritdoc/>
        public int         Count              => _list.Count;
        /// <inheritdoc/>
        public string      this[string Index] => _list.FirstOrDefault(i => i.Key == Index).Value
                                              ?? Index;
    }
}
