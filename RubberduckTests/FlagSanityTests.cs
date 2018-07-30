using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.WindowsApi;

namespace RubberduckTests
{
    [TestFixture]
    public class FlagSanityTests
    {
        private static List<object> TypedEnumValues<T>() 
        {
            return Enum.GetValues(typeof(T))
                .Cast<T>()
                .Select(member => Convert.ChangeType(member, Enum.GetUnderlyingType(typeof(T))))
                .ToList();
        }

        [Test]
        [Category("Flag Enumerations")]
        public void VBENativeServices_WindowType_HasNoOverlap()
        {
            var values = TypedEnumValues<WindowType>();
            var distinct = values.Distinct().ToList();

            Assert.AreEqual(values.Count, distinct.Count);
        }

        [Test]
        [Category("Flag Enumerations")]
        public void DeclarationType_HasNoOverlap()
        {
            var values = TypedEnumValues<DeclarationType>();
            var distinct = values.Distinct().ToList();

            Assert.AreEqual(values.Count, distinct.Count);
        }
    }
}
