using Rubberduck.Resources.Registration;
using Rubberduck.VBEditor;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    [ComVisible(true)]
    [Guid(RubberduckGuid.ITimesGuid)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITimes
    {
        [Description("Specifies that a mocked method should be invoked CallCount times as maximum.")]
        ITimes AtMost(int CallCount);

        [Description("Specifies that a mocked method should be invoked one time as maximum.")]
        ITimes AtMostOnce();

        [Description("Specifies that a mocked method should be invoked CallCount times as minimum.")]
        ITimes AtLeast(int CallCount);

        [Description("Specifies that a mocked method should be invoked one time as minimum.")]
        ITimes AtLeastOnce();

        [Description("Specifies that a mocked method should be invoked between MinCallCount and MaxCallCount times.")]
        ITimes Between(int MinCallCount, int MaxCallCount, SetupArgumentRange RangeKind = SetupArgumentRange.Inclusive);

        [Description("Specifies that a mocked method should be invoked exactly CallCount times.")]
        ITimes Exactly(int CallCount);

        [Description("Specifies that a mocked method should be invoked exactly one time.")]
        ITimes Once();

        [Description("Specifies that a mocked method should not be invoked.")]
        ITimes Never();
    }


    [ComVisible(true)]
    [Guid(RubberduckGuid.TimesGuid)]
    [ProgId(RubberduckProgId.TimesProgId)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITimes))]
    public class Times : ITimes, IEquatable<Times>
    {
        internal int Min { get; }
        internal int Max { get; }

        internal Times() { }

        internal Times(int min, int max)
        {
            Min = min;
            Max = max;
        }

        public bool Equals(Times other) => other.Min == Min && other.Max == Max;

        public override bool Equals(object obj) => Equals((Times)obj);

        public override int GetHashCode() => HashCode.Compute(Min, Max);

        public ITimes AtMost(int CallCount) => new Times(0, CallCount);

        public ITimes AtMostOnce() => new Times(0, 1);

        public ITimes AtLeast(int CallCount) => new Times(CallCount, int.MaxValue);

        public ITimes AtLeastOnce() => new Times(1, int.MaxValue);

        public ITimes Between(int MinCallCount, int MaxCallCount, SetupArgumentRange RangeKind = SetupArgumentRange.Inclusive)
            => new Times(MinCallCount + RangeKind == SetupArgumentRange.Exclusive ? 1 : 0, MaxCallCount - RangeKind == SetupArgumentRange.Exclusive ? 1 : 0);

        public ITimes Exactly(int CallCount) => new Times(CallCount, CallCount);

        public ITimes Once() => new Times(1, 1);

        public ITimes Never() => new Times(0, 0);

        public static bool operator ==(Times lhs, Times rhs) => lhs.Equals(rhs);
        public static bool operator !=(Times lhs, Times rhs) => !lhs.Equals(rhs);

        public void Deconstruct(out int MinCallCount, out int MaxCallCount)
        {
            MinCallCount = Min;
            MaxCallCount = Max;
        }
    }
}
