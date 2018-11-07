using Rubberduck.AutoComplete;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    [Flags]
    public enum ModifierKeySetting
    {
        None = 0,
        CtrlKey = 1,
        ShiftKey = 2,
    }

    public interface IAutoCompleteSettings
    {
        bool IsEnabled { get; set; }
        AutoCompleteSettings.SmartConcatSettings SmartConcat { get; set; }
        AutoCompleteSettings.SelfClosingPairSettings SelfClosingPairs { get; set; }
        AutoCompleteSettings.BlockCompletionSettings BlockCompletion { get; set; }
    }

    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlType(AnonymousType = true)]
    public class AutoCompleteSettings : IAutoCompleteSettings, IEquatable<AutoCompleteSettings>
    {
        public static AutoCompleteSettings AllEnabled =>
            new AutoCompleteSettings
            {
                IsEnabled = true,
                BlockCompletion =
                    new BlockCompletionSettings {IsEnabled = true, CompleteOnEnter = true, CompleteOnTab = true},
                SmartConcat =
                    new SmartConcatSettings {IsEnabled = true, ConcatVbNewLineModifier = ModifierKeySetting.CtrlKey},
                SelfClosingPairs = 
                    new SelfClosingPairSettings {IsEnabled = true}
            };

        public AutoCompleteSettings()
        {
            SmartConcat = new SmartConcatSettings();
            SelfClosingPairs = new SelfClosingPairSettings();
            BlockCompletion = new BlockCompletionSettings();
        }

        [XmlAttribute]
        public bool IsEnabled { get; set; }

        public SmartConcatSettings SmartConcat { get; set; }

        public SelfClosingPairSettings SelfClosingPairs { get; set; }

        public BlockCompletionSettings BlockCompletion { get; set; }

        public class SmartConcatSettings : IEquatable<SmartConcatSettings>
        {
            public bool IsEnabled { get; set; }
            public ModifierKeySetting ConcatVbNewLineModifier { get; set; }

            public bool Equals(SmartConcatSettings other)
                => other != null &&
                   other.IsEnabled == IsEnabled &&
                   other.ConcatVbNewLineModifier == ConcatVbNewLineModifier;
        }

        public class SelfClosingPairSettings : IEquatable<SelfClosingPairSettings>
        {
            [XmlAttribute]
            public bool IsEnabled { get; set; }

            public bool Equals(SelfClosingPairSettings other)
                => other != null &&
                   other.IsEnabled == IsEnabled;
        }

        public class BlockCompletionSettings : IEquatable<BlockCompletionSettings>
        {
            [XmlAttribute]
            public bool IsEnabled { get; set; }
            [XmlAttribute]
            public bool CompleteOnEnter { get; set; }
            [XmlAttribute]
            public bool CompleteOnTab { get; set; }

            public bool Equals(BlockCompletionSettings other)
                => other != null &&
                   other.IsEnabled == IsEnabled &&
                   other.CompleteOnEnter == CompleteOnEnter &&
                   other.CompleteOnTab == CompleteOnTab;
        }

        public bool Equals(AutoCompleteSettings other)
            => other != null &&
               other.IsEnabled == IsEnabled &&
               other.BlockCompletion.Equals(BlockCompletion) &&
               other.SmartConcat.Equals(SmartConcat) &&
               other.SelfClosingPairs.Equals(SelfClosingPairs);
    }
}
