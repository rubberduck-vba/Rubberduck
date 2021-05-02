using System.Runtime.InteropServices;
using Moq;
using Rubberduck.Resources.Registration;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.SetupArgumentRangeGuid)
    ]
    public enum SetupArgumentRange
    {
        Inclusive = Range.Inclusive,
        Exclusive = Range.Exclusive
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.SetupArgumentTypeGuid)
    ]
    public enum SetupArgumentType
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
        Guid(RubberduckGuid.ISetupArgumentDefinitionGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface ISetupArgumentDefinition
    {
        [DispId(1)]
        SetupArgumentType Type { get; }
        
        [DispId(2)]
        object[] Values { get; }

        [DispId(3)]
        SetupArgumentRange? Range
        {
            [return: MarshalAs(UnmanagedType.Struct)]
            get;
        }
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.SetupArgumentDefinitionGuid),
        ProgId(RubberduckProgId.SetupArgumentDefinitionProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(ISetupArgumentDefinition))
    ]
    public class SetupArgumentDefinition : ISetupArgumentDefinition
    {
        internal static SetupArgumentDefinition CreateIs(object value)
        {
            return new SetupArgumentDefinition(SetupArgumentType.Is, new[] { value });
        }

        internal static SetupArgumentDefinition CreateIsAny()
        {
            return new SetupArgumentDefinition(SetupArgumentType.IsAny, null);
        }

        internal static SetupArgumentDefinition CreateIsIn(object[] values)
        {
            return new SetupArgumentDefinition(SetupArgumentType.IsIn, values);
        }

        internal static SetupArgumentDefinition CreateIsInRange(object start, object end, SetupArgumentRange range)
        {
            return new SetupArgumentDefinition(SetupArgumentType.IsInRange, new[] { start, end }, range);
        }

        internal static SetupArgumentDefinition CreateIsNotIn(object[] values)
        {
            return new SetupArgumentDefinition(SetupArgumentType.IsNotIn, values);
        }

        internal static SetupArgumentDefinition CreateIsNotNull()
        {
            return new SetupArgumentDefinition(SetupArgumentType.IsNotNull, null);
        }

        private SetupArgumentDefinition(SetupArgumentType type, object[] values)
        {
            Type = type;
            Values = values;
        }

        private SetupArgumentDefinition(SetupArgumentType type, object[] values, SetupArgumentRange range) : 
            this(type, values)
        {
            Range = range;
        }

        public SetupArgumentType Type { get; }
        public object[] Values { get; }
        public SetupArgumentRange? Range { get; }
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.ISetupArgumentCreatorGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface ISetupArgumentCreator
    {
        SetupArgumentDefinition Is(object Value);
        SetupArgumentDefinition IsAny();
        SetupArgumentDefinition IsIn(object[] Values);
        SetupArgumentDefinition IsInRange(object Start, object End, SetupArgumentRange Range);
        SetupArgumentDefinition IsNotIn(object[] Values);
        SetupArgumentDefinition IsNotNull();
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.SetupArgumentCreatorGuid),
        ProgId(RubberduckProgId.SetupArgumentCreatorProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(ISetupArgumentCreator))
    ]
    public class SetupArgumentCreator : ISetupArgumentCreator
    {
        public SetupArgumentDefinition Is(object Value)
        {
            return SetupArgumentDefinition.CreateIs(Value);
        }

        public SetupArgumentDefinition IsAny()
        {
            return SetupArgumentDefinition.CreateIsAny();
        }

        public SetupArgumentDefinition IsIn(object[] Values)
        {
            return SetupArgumentDefinition.CreateIsIn(Values);
        }

        public SetupArgumentDefinition IsInRange(object Start, object End, SetupArgumentRange Range)
        {
            return SetupArgumentDefinition.CreateIsInRange(Start, End, Range);
        }

        public SetupArgumentDefinition IsNotIn(object[] Values)
        {
            return SetupArgumentDefinition.CreateIsNotIn(Values);
        }

        public SetupArgumentDefinition IsNotNull()
        {
            return SetupArgumentDefinition.CreateIsNotNull();
        }
    }
}
