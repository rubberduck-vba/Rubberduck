using System;
using System.Reflection;
using ExpressiveReflection;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    public class ItByRefMemberInfos
    {
        public static Type ItByRef(Type type)
        {
            return typeof(ItByRef<>).MakeGenericType(type);
        }

        public static MethodInfo Is(Type type)
        {
            return Reflection.GetMethodExt(ItByRef(type), nameof(ItByRef<object>.Is), type);
        }

        public static FieldInfo Value(Type type)
        {
            return ItByRef(type).GetField(nameof(ItByRef<object>.Value));
        }
    }
}
