using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Reference : SafeComWrapper<VB.Reference>, IReference
    {
        public Reference(VB.Reference target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public string Name
        {
            get { return IsBroken ? string.Empty : Target.Name; }
        }

        public string Guid
        {
            get { return IsBroken ? string.Empty : Target.Guid; }
        }

        public int Major
        {
            get { return IsBroken ? 0 : Target.Major; }
        }

        public int Minor
        {
            get { return IsBroken ? 0 : Target.Minor; }
        }

        public string Version
        {
            get { return string.Format($"{Major}.{Minor}"); }
        }

        public string Description
        {
            get { return IsBroken ? string.Empty : Target.Description; }
        }

        public string FullPath
        {
            get
            {
                try
                {
                    return IsBroken ? string.Empty : Target.FullPath;
                }
                catch (COMException)
                {
                    return string.Empty;
                }
            }
        }

        public bool IsBuiltIn
        {
            get { return !IsBroken && Target.BuiltIn; }
        }

        public bool IsBroken
        {
            get
            {
                try
                {
                    return IsWrappingNullReference || Target.IsBroken;
                }
                catch (COMException ex)
                {
                    _logger.Trace(ex, "IsBroken is broken.");
                }

                return true;
            }
        }

        public ReferenceKind Type
        {
            get { return IsBroken ? 0 : (ReferenceKind)Target.Type; }
        }

        public IReferences Collection
        {
            get { return new References(IsBroken ? null : Target.Collection); }
        }

        public IVBE VBE
        {
            get { return new VBE(IsBroken ? null : Target.VBE); }
        }

        public override bool Equals(ISafeComWrapper<VB.Reference> other)
        {
            return IsEqualIfNull(other) ||
                   (other != null 
                    && (int)other.Target.Type == (int)Type
                    && other.Target.Name == Name
                    && other.Target.Guid == Guid
                    && other.Target.FullPath == FullPath
                    && other.Target.Major == Major
                    && other.Target.Minor == Minor);
        }

        public bool Equals(IReference other)
        {
            return Equals(other as SafeComWrapper<VB.Reference>);
        }

        public override int GetHashCode()
        {
            return IsBroken ? 0 : HashCode.Compute(Type, Name, Guid, FullPath, Major, Minor);
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}