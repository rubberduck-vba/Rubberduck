using System.Runtime.InteropServices;
using Moq;
using Rubberduck.Resources.Registration;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.MockArgumentRangeGuid)
    ]
    public enum MockArgumentRange
    {
        Inclusive = Range.Inclusive,
        Exclusive = Range.Exclusive
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.MockArgumentTypeGuid)
    ]
    public enum MockArgumentType
    {
        Is,
        IsAny,
        IsIn,
        IsInRange,
        IsNotIn,
        IsNotNull
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.IMockArgumentDefinitionGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IMockArgumentDefinition
    {
        [DispId(1)]
        MockArgumentType Type { get; }
        
        [DispId(2)]
        object[] Values { get; }

        [DispId(3)]
        MockArgumentRange? Range
        {
            [return: MarshalAs(UnmanagedType.Struct)]
            get;
        }
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.MockArgumentDefinitionGuid),
        ProgId(RubberduckProgId.MockArgumentDefinitionProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IMockArgumentDefinition))
    ]
    public class MockArgumentDefinition : IMockArgumentDefinition
    {
        internal static MockArgumentDefinition CreateIs(object value)
        {
            return new MockArgumentDefinition(MockArgumentType.Is, new[] { value });
        }

        internal static MockArgumentDefinition CreateIsAny()
        {
            return new MockArgumentDefinition(MockArgumentType.IsAny, null);
        }

        internal static MockArgumentDefinition CreateIsIn(object[] values)
        {
            return new MockArgumentDefinition(MockArgumentType.IsIn, values);
        }

        internal static MockArgumentDefinition CreateIsInRange(object start, object end, MockArgumentRange range)
        {
            return new MockArgumentDefinition(MockArgumentType.IsInRange, new[] { start, end }, range);
        }

        internal static MockArgumentDefinition CreateIsNotIn(object[] values)
        {
            return new MockArgumentDefinition(MockArgumentType.IsNotIn, values);
        }

        internal static MockArgumentDefinition CreateIsNotNull()
        {
            return new MockArgumentDefinition(MockArgumentType.IsNotNull, null);
        }

        private MockArgumentDefinition(MockArgumentType type, object[] values)
        {
            Type = type;
            Values = values;
        }

        private MockArgumentDefinition(MockArgumentType type, object[] values, MockArgumentRange range) : 
            this(type, values)
        {
            Range = range;
        }

        public MockArgumentType Type { get; }
        public object[] Values { get; }
        public MockArgumentRange? Range { get; }
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.IMockArgumentCreatorGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IMockArgumentCreator
    {
        MockArgumentDefinition Is(object Value);
        MockArgumentDefinition IsAny();
        MockArgumentDefinition IsIn(object[] Values);
        MockArgumentDefinition IsInRange(object Start, object End, MockArgumentRange Range);
        MockArgumentDefinition IsNotIn(object[] Values);
        MockArgumentDefinition IsNotNull();
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.MockArgumentCreatorGuid),
        ProgId(RubberduckProgId.MockArgumentCreatorProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IMockArgumentCreator))
    ]
    public class MockArgumentCreator : IMockArgumentCreator
    {
        public MockArgumentDefinition Is(object Value)
        {
            return MockArgumentDefinition.CreateIs(Value);
        }

        public MockArgumentDefinition IsAny()
        {
            return MockArgumentDefinition.CreateIsAny();
        }

        public MockArgumentDefinition IsIn(object[] Values)
        {
            return MockArgumentDefinition.CreateIsIn(Values);
        }

        public MockArgumentDefinition IsInRange(object Start, object End, MockArgumentRange Range)
        {
            return MockArgumentDefinition.CreateIsInRange(Start, End, Range);
        }

        public MockArgumentDefinition IsNotIn(object[] Values)
        {
            return MockArgumentDefinition.CreateIsNotIn(Values);
        }

        public MockArgumentDefinition IsNotNull()
        {
            return MockArgumentDefinition.CreateIsNotNull();
        }
    }
}
