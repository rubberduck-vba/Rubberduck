using System.Collections.Generic;
using System.Dynamic;

namespace Rubberduck.Parsing.Inspections
{
    public class PropertyBag : DynamicObject
    {
        private readonly Dictionary<string, object> _properties = new Dictionary<string, object>();

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            return _properties.TryGetValue(binder.Name, out result);
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            _properties[binder.Name] = value;
            return true;
        }

        public override IEnumerable<string> GetDynamicMemberNames()
        {
            return _properties.Keys;
        }
    }
}
