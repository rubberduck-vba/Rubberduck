using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IFieldEncapsulationAttributes
    {
        string FieldName { get; set; }
        string NewFieldName { get; set; }
        string PropertyName { get; set; }
        string AsTypeName { get; set; }
        string ParameterName { get; set; }
        bool ImplementLetSetterType { get; set; }
        bool ImplementSetSetterType { get; set; }
        bool IsFlaggedToEncapsulate { get; set; }
        bool IsValueType { get; set; }
        bool IsVariantType { get; set; }
        bool IsUserDefinedType { set; get; }
        bool IsArray { set; get; }
        bool IsObjectType  { set;  get; }
        bool CanImplementLet { get; set; }
        bool CanImplementSet { get; set; }
    }

    public interface IUDTFieldEncapsulationAttributes : IFieldEncapsulationAttributes
    {
        bool MemberEncapsulationFlag(string name);
        bool EncapsulateAllUDTMembers { get; set; }
        IEnumerable<(string Name, bool Encapsulate)> MemberFlags { get; }
    }

    public class FieldEncapsulationAttributes : IFieldEncapsulationAttributes
    {
        private static string DEFAULT_LET_PARAMETER = "value";

        //Only used by tests
        public FieldEncapsulationAttributes() { }

        public FieldEncapsulationAttributes(Declaration target, string newFieldName = null, string parameterName = null)
        {
            FieldName = target.IdentifierName;
            NewFieldName = newFieldName ?? target.IdentifierName;
            PropertyName = target.IdentifierName;
            AsTypeName = target.AsTypeName;
            ParameterName = parameterName ?? DEFAULT_LET_PARAMETER;
            IsFlaggedToEncapsulate = false;
            IsVariantType = target.AsTypeName?.Equals(Tokens.Variant) ?? true;
            IsValueType = !IsVariantType && (SymbolList.ValueTypes.Contains(target.AsTypeName) ||
                                             target.DeclarationType == DeclarationType.Enumeration);
            IsObjectType = target.IsObject;
            IsUserDefinedType = false;
            IsArray = target.IsArray;
            ImplementLetSetterType = UpdateFieldUsingLet();
            ImplementSetSetterType = UpdateFieldUsingSet();
            CanImplementLet = false;
            CanImplementSet = false;
        }

        public FieldEncapsulationAttributes(IFieldEncapsulationAttributes attributes)
        {
            FieldName = attributes.FieldName;
            NewFieldName = attributes.NewFieldName ?? attributes.FieldName;
            PropertyName = attributes.PropertyName;
            AsTypeName = attributes.AsTypeName;
            ParameterName = attributes.ParameterName;
            IsFlaggedToEncapsulate = attributes.IsFlaggedToEncapsulate;
            IsVariantType = attributes.IsVariantType;
            IsValueType = attributes.IsValueType;
            IsUserDefinedType = attributes.IsUserDefinedType;
            IsArray = attributes.IsArray;
            IsObjectType = attributes.IsObjectType;
            ImplementLetSetterType = attributes.ImplementLetSetterType;
            ImplementSetSetterType = attributes.ImplementSetSetterType;
        }

        public string FieldName { get; set; }

        private string _newFieldName;
        public string NewFieldName
        {
            get => _newFieldName ?? FieldName;
            set => _newFieldName = value;
        }
        public string PropertyName { get; set; }
        public string AsTypeName { get; set; }
        public string ParameterName { get; set; } = DEFAULT_LET_PARAMETER;

        private bool? _implLet;
        public bool ImplementLetSetterType
        {
            get => _implLet.HasValue ? _implLet.Value : UpdateFieldUsingLet();
            set => _implLet = value;
        }

        private bool? _implSet;
        public bool ImplementSetSetterType
        {
            get => _implSet.HasValue ? _implSet.Value : UpdateFieldUsingSet();
            set => _implSet = value;
        }

        public bool IsFlaggedToEncapsulate { get; set; }
        public bool IsValueType { set; get; }
        public bool IsVariantType { set; get; }
        public bool IsUserDefinedType { set; get; }
        public bool IsArray { set; get; }

        private bool? _isObjectType;
        public bool IsObjectType
        {
            get => _isObjectType.HasValue ? _isObjectType.Value : !(IsValueType || IsVariantType);
            set => _isObjectType = value;
        }
        public bool CanImplementLet { get; set; }
        public bool CanImplementSet { get; set; }

        private bool UpdateFieldUsingSet()
            => IsObjectType || !IsArray && !IsUserDefinedType && !IsValueType && IsVariantType;

        private bool UpdateFieldUsingLet()
            => !IsObjectType && !IsArray && (IsUserDefinedType || IsValueType || IsVariantType);
    }

    public class UDTFieldEncapsulationAttributes : IUDTFieldEncapsulationAttributes
    {
        public UDTFieldEncapsulationAttributes(FieldEncapsulationAttributes attributes, IEnumerable<string> udtMemberNames)
        {
            _attributes = attributes;
            EncapsulateAllUDTMembers = false;
            IsVariantType = false;
            IsValueType = false;
            IsUserDefinedType = true;
            IsArray = false;

            foreach (var udtMemberName in udtMemberNames)
            {
                if (!_memberEncapsulationFlags.ContainsKey(udtMemberName))
                {
                    _memberEncapsulationFlags.Add(udtMemberName, false);
                }
            }
        }

        //TODO: Copy ctor may not every be needed
        public UDTFieldEncapsulationAttributes(IUDTFieldEncapsulationAttributes attributes)
        {
            FieldName = attributes.FieldName;
            NewFieldName = attributes.NewFieldName ?? attributes.FieldName;
            PropertyName = attributes.PropertyName;
            AsTypeName = attributes.AsTypeName;
            ParameterName = attributes.ParameterName;
            IsFlaggedToEncapsulate = attributes.IsFlaggedToEncapsulate;
            IsVariantType = false;
            IsValueType = false;
            IsUserDefinedType = true;
            IsArray = false;
            ImplementLetSetterType = attributes.ImplementLetSetterType;
            ImplementSetSetterType = attributes.ImplementSetSetterType;

            //TODO: Dumb to have a dictionary here - it is a tuple at the moment(!?)
            _memberEncapsulationFlags = attributes.MemberFlags.ToDictionary(k => k.Name, e => e.Encapsulate);
        }

        private IFieldEncapsulationAttributes _attributes { set; get; }
        private Dictionary<string, bool> _memberEncapsulationFlags = new Dictionary<string, bool>();
        
        //Only used by tests
        public void FlagUdtMemberEncapsulation(bool encapsulateFlag, params string[] names)
        {
            foreach (var name in names)
            {
                if (_memberEncapsulationFlags.ContainsKey(name))
                {
                    _memberEncapsulationFlags[name] = encapsulateFlag;
                }
                else
                {
                    _memberEncapsulationFlags.Add(name, encapsulateFlag);
                }
            }
        }

        public bool EncapsulateAllUDTMembers
        {
            set => SetEncapsulationFlagForAllMembers(value);
            get => _memberEncapsulationFlags.Values.All(v => v == true);
        }

        public string FieldName
        {
            get => _attributes.FieldName;
            set => _attributes.FieldName = value;
        }

        public string PropertyName
        {
            get => _attributes.PropertyName;
            set => _attributes.PropertyName = value;
        }

        public string NewFieldName
        {
            get => _attributes.NewFieldName;
            set => _attributes.NewFieldName = value;
        }

        public string AsTypeName
        {
            get => _attributes.AsTypeName;
            set => _attributes.AsTypeName = value;
        }
        public string ParameterName
        {
            get => _attributes.ParameterName;
            set => _attributes.ParameterName = value;
        }
        public bool ImplementLetSetterType
        {
            get => _attributes.ImplementLetSetterType;
            set => _attributes.ImplementLetSetterType = value;
        }

        public bool ImplementSetSetterType
        {
            get => _attributes.ImplementSetSetterType;
            set => _attributes.ImplementSetSetterType = value;
        }

        public bool IsFlaggedToEncapsulate
        {
            get => _attributes.IsFlaggedToEncapsulate;
            set => _attributes.IsFlaggedToEncapsulate = value;
        }

        public bool IsValueType
        {
            get => _attributes.IsValueType;
            set => _attributes.IsValueType = value;
        }

        public bool IsVariantType
        {
            get => _attributes.IsVariantType;
            set => _attributes.IsVariantType = value;
        }

        public bool IsUserDefinedType
        {
            get => _attributes.IsUserDefinedType;
            set => _attributes.IsUserDefinedType = value;
        }

        public bool IsArray
        {
            get => _attributes.IsArray;
            set => _attributes.IsArray = value;
        }

        public bool IsObjectType
        {
            get => _attributes.IsObjectType;
            set => _attributes.IsObjectType = value;
        }

        public bool CanImplementLet { get; set; }
        public bool CanImplementSet { get; set; }

        private void SetEncapsulationFlagForAllMembers(bool flag)
        {
            foreach (var key in _memberEncapsulationFlags.Keys.ToList())
            {
                _memberEncapsulationFlags[key] = flag;
            }
        }

        public IEnumerable<(string Name, bool Encapsulate)> MemberFlags
        {
            get
            {
                var results = new List<(string Name, bool Encapsulate)>();
                foreach (var name in _memberEncapsulationFlags.Keys)
                {
                    results.Add((name, _memberEncapsulationFlags[name]));
                }
                return results;
            }
        } 
             
        public bool MemberEncapsulationFlag(string name)
        {
            return _memberEncapsulationFlags.ContainsKey(name) ? _memberEncapsulationFlags[name] : false;
        }
    }

    //public class UDTMemberEncapsulationAttributes : IFieldEncapsulationAttributes
    //{
    //    public UDTMemberEncapsulationAttributes(IFieldEncapsulationAttributes attributes)
    //    {
    //        _attributes = attributes;
    //    }

    //    private IFieldEncapsulationAttributes _attributes { set; get; }

    //    public string FieldName
    //    {
    //        get => _attributes.FieldName;
    //        set => _attributes.FieldName = value;
    //    }

    //    public string PropertyName
    //    {
    //        get => _attributes.PropertyName;
    //        set => _attributes.PropertyName = value;
    //    }

    //    public string NewFieldName
    //    {
    //        get => _attributes.NewFieldName;
    //        set => _attributes.NewFieldName = value;
    //    }

    //    public string AsTypeName
    //    {
    //        get => _attributes.AsTypeName;
    //        set => _attributes.AsTypeName = value;
    //    }
    //    public string ParameterName
    //    {
    //        get => _attributes.ParameterName;
    //        set => _attributes.ParameterName = value;
    //    }
    //    public bool ImplementLetSetterType
    //    {
    //        get => _attributes.ImplementLetSetterType;
    //        set => _attributes.ImplementLetSetterType = value;
    //    }

    //    public bool ImplementSetSetterType
    //    {
    //        get => _attributes.ImplementSetSetterType;
    //        set => _attributes.ImplementSetSetterType = value;
    //    }
    //    public bool IsFlaggedToEncapsulate
    //    {
    //        get => _attributes.IsFlaggedToEncapsulate;
    //        set => _attributes.IsFlaggedToEncapsulate = value;
    //    }

    //    public bool IsValueType
    //    {
    //        get => _attributes.IsValueType;
    //        set => _attributes.IsValueType = value;
    //    }

    //    public bool IsVariantType
    //    {
    //        get => _attributes.IsVariantType;
    //        set => _attributes.IsVariantType = value;
    //    }

    //    public bool IsUserDefinedType
    //    {
    //        get => _attributes.IsUserDefinedType;
    //        set => _attributes.IsUserDefinedType = value;
    //    }

    //    public bool IsObjectType
    //    {
    //        get => _attributes.IsObjectType;
    //        set => _attributes.IsObjectType = value;
    //    }
    //    //public bool IsObjectType => _attributes.IsObjectType;
    //}
}
