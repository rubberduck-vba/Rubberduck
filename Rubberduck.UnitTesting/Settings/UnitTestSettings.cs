using System;
using System.Configuration;
using System.Xml.Serialization;

namespace Rubberduck.UnitTesting.Settings
{
    public enum BindingMode
    {
        EarlyBinding,
        LateBinding,
        DualBinding
    }

    public enum AssertMode
    {
        StrictAssert,
        PermissiveAssert
    }

    public interface IUnitTestSettings
    {
        BindingMode BindingMode { get; set; }
        AssertMode AssertMode { get; set; }

        bool ModuleInit { get; set; }
        bool MethodInit { get; set; }
        bool DefaultTestStubInNewModule { get; set; }
    }

    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlType(AnonymousType = true)]
    public class UnitTestSettings : IUnitTestSettings, IEquatable<UnitTestSettings>
    {
        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public UnitTestSettings()
        {
        }

        public UnitTestSettings(BindingMode bindingMode, AssertMode assertMode, bool moduleInit, bool methodInit, bool defaultTestStub)
        {
            BindingMode = bindingMode;
            AssertMode = assertMode;
            ModuleInit = moduleInit;
            MethodInit = methodInit;
            DefaultTestStubInNewModule = defaultTestStub;
        }

        public BindingMode BindingMode { get; set; }
        public AssertMode AssertMode { get; set; }
        public bool ModuleInit { get; set; }
        public bool MethodInit { get; set; }
        public bool DefaultTestStubInNewModule { get; set; }

        public bool Equals(UnitTestSettings other)
        {
            return other != null &&
                   BindingMode == other.BindingMode &&
                   AssertMode == other.AssertMode &&
                   ModuleInit == other.ModuleInit &&
                   MethodInit == other.MethodInit &&
                   DefaultTestStubInNewModule == other.DefaultTestStubInNewModule;
        }
    }
}
