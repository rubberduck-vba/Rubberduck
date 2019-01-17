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
        internal MockArgumentDefinition(MockArgumentType type, object[] values)
        {
            Type = type;
            Values = values;
        }

        internal MockArgumentDefinition(MockArgumentType type, object[] values, MockArgumentRange range) : 
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
            return new MockArgumentDefinition(MockArgumentType.Is, new[] {Value});
        }

        public MockArgumentDefinition IsAny()
        {
            return new MockArgumentDefinition(MockArgumentType.IsAny, null);
        }

        public MockArgumentDefinition IsIn(object[] Values)
        {
            return new MockArgumentDefinition(MockArgumentType.IsIn, Values);
        }

        public MockArgumentDefinition IsInRange(object Start, object End, MockArgumentRange Range)
        {
            return new MockArgumentDefinition(MockArgumentType.IsInRange, new[] {Start, End}, Range);
        }

        public MockArgumentDefinition IsNotIn(object[] Values)
        {
            return new MockArgumentDefinition(MockArgumentType.IsNotIn, Values);
        }

        public MockArgumentDefinition IsNotNull()
        {
            return new MockArgumentDefinition(MockArgumentType.IsNotNull, null);
        }
    }
}
