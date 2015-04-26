using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Parsing
{
    public struct QualifiedModuleName
    {
        public QualifiedModuleName(VBComponent component)
        {
            _component = component;
            _componentName = component == null ? string.Empty : component.Name;
            _projectName = component == null ? string.Empty : component.Collection.Parent.Name;
            _projectHashCode = component == null ? 0 : component.Collection.Parent.GetHashCode();

            var module = _component.CodeModule;
            _contentHashCode = module.CountOfLines > 0 
                ? module.get_Lines(1, module.CountOfLines).GetHashCode() 
                : 0;
        }

        public QualifiedMemberName QualifyMemberName(string member)
        {
            return new QualifiedMemberName(this, member);
        }

        private readonly VBComponent _component;
        public VBComponent Component { get { return _component; } }
        public VBProject Project { get { return _component == null ? null : _component.Collection.Parent; } }

        private readonly int _projectHashCode;
        public int ProjectHashCode { get { return _projectHashCode; } }

        private readonly int _contentHashCode;

        private readonly string _projectName;
        public string ProjectName { get { return _projectName;} }

        private readonly string _componentName;
        public string ComponentName { get { return _componentName; } }

        public override string ToString()
        {
            return _component == null ? string.Empty : _projectName + "." + _componentName;
        }

        public override int GetHashCode()
        {
            return _component == null ? 0 : _component.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            try
            {
                var other = (QualifiedModuleName)obj;
                return other.Component == Component && other._contentHashCode == _contentHashCode;
            }
            catch (InvalidCastException)
            {
                return false;
            }
        }

        public static bool operator ==(QualifiedModuleName a, QualifiedModuleName b)
        {
            return a.Equals(b);
        }

        public static bool operator !=(QualifiedModuleName a, QualifiedModuleName b)
        {
            return !a.Equals(b);
        }
    }
}