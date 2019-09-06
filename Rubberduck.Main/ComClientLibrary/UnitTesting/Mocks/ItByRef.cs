namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    public class ItByRef<T>
    {
        // Only a field of a class can be passed by-ref.
        public T Value;

        public delegate void ByRefCallback(ref T input);

        public ByRefCallback Callback { get; }

        public static ItByRef<T> Is(T initialValue)
        {
            return Is(initialValue, null);
        }

        public static ItByRef<T> Is(T initialValue, ByRefCallback action)
        {
            return new ItByRef<T>(initialValue, action);
        }

        private ItByRef(T initialValue, ByRefCallback callback)
        {
            Value = initialValue;
            Callback = callback;
        }
    }
}
