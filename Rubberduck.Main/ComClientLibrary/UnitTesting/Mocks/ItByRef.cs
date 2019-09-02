using System;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    public class ItByRef<T>
    {
        // Only a field of a class can be passed by-ref.
        public T Value;

        public static ItByRef<T> Is(T input, Func<T, bool> condition)
        {
            return condition(input) 
                ? new ItByRef<T>(input) 
                : new ItByRef<T>(default);
        }

        private ItByRef(T input)
        {
            Value = input;
        }
    }
}
