using System;

namespace Rubberduck.UnitTesting
{
    public class TestCategory
    {
        public string Name { get; }

        public TestCategory(string name)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
        }

        public override bool Equals(object obj)
        {
            if (!(obj is TestCategory category))
            {
                return false;
            }

            return Name == category.Name;
        }

        public override int GetHashCode()
        {
            return Name.GetHashCode();
        }
    }
}
