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


    public static class MoqTimesExt
    {
        public static ITimes ToRubberduckTimes(this Moq.Times times) => new Times(times);
    }

    [ComVisible(true)]
    [Guid(RubberduckGuid.TimesGuid)]
    [ProgId(RubberduckProgId.TimesProgId)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITimes))]
    public class Times : ITimes, IEquatable<Times>
    {
        internal Times() { }

        internal Times(Moq.Times moqTimes)
        {
            MoqTimes = moqTimes;
        }

        public bool Equals(Times other) => MoqTimes.Equals(other?.MoqTimes);

        public override bool Equals(object obj) => Equals(obj as Times);

        public override int GetHashCode() => MoqTimes.GetHashCode();

        public ITimes AtMost(int CallCount) => new Times(Moq.Times.AtMost(CallCount));

        public ITimes AtMostOnce() => new Times(Moq.Times.AtMostOnce());

        public ITimes AtLeast(int CallCount) => new Times(Moq.Times.AtLeast(CallCount));

        public ITimes AtLeastOnce() => new Times(Moq.Times.AtLeastOnce());

        public ITimes Between(int MinCallCount, int MaxCallCount, SetupArgumentRange RangeKind = SetupArgumentRange.Inclusive)
            => new Times(Moq.Times.Between(MinCallCount, MaxCallCount, RangeKind == SetupArgumentRange.Exclusive ? Moq.Range.Exclusive : Moq.Range.Inclusive ));

        public ITimes Exactly(int CallCount) => new Times(Moq.Times.Exactly(CallCount));

        public ITimes Once() => new Times(Moq.Times.Once());

        public ITimes Never() => new Times(Moq.Times.Never());

        internal Moq.Times MoqTimes { get; }
    }
}
