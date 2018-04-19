using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VBControl : SafeComWrapper<VB.VBControl>, IControl
    {
        public VBControl(VB.VBControl target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public IProperties Properties => new Properties(IsWrappingNullReference? null : Target.Properties);
        
        public string Name
        {
            get
            {
                if (IsWrappingNullReference)
                {
                    return string.Empty;
                }

                using (var properties = this.Properties)
                {
                    foreach (var property in properties)
                    {
                        using (property)
                        {
                            if (property.Name == "Name")
                            {
                                return (string)property.Value;
                            }
                        }
                    }
                }
                // Should never happen - all VB controls are required to be named.
                return null;
            }
            set
            {
                if (!IsWrappingNullReference)
                {
                    using (var properties = this.Properties)
                    {
                        foreach (var property in properties)
                        {
                            using (property)
                            {
                                if (property.Name == "Name")
                                {
                                    property.Value = value;
                                    return;
                                }
                            }
                        }                        
                    }
                    // Should never happen - all VB controls are required to be named.                    
                }                
            }
        }

        public override bool Equals(ISafeComWrapper<VB.VBControl> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IControl other)
        {
            return Equals(other as SafeComWrapper<VB.VBControl>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}