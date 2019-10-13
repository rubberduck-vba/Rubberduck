using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization;
using Rubberduck.Parsing.Symbols;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using FUNCFLAGS = System.Runtime.InteropServices.ComTypes.FUNCFLAGS;

namespace Rubberduck.Parsing.ComReflection
{
    internal enum DispId
    {
        Collect = -8,           //The method you are calling through Invoke is an accessor function. 
        Destructor = -7,        //The C++ destructor function for the object. 
        Construtor = -6,        //The C++ constructor function for the object. 
        Evaluate = -5,          //This method is implicitly invoked when the ActiveX client encloses the arguments in square brackets.
        NewEnum = -4,           //It returns an enumerator object that supports IEnumVARIANT.
        PropertyPut = -3,       //The parameter that receives the value of an assignment in a PROPERTYPUT. 
        Unknown = -1,           //The value returned by IDispatch::GetIDsOfNames to indicate that a member or parameter name was not found.
        Value = 0               //The default member for the object.
    }

    [DataContract]
    [KnownType(typeof(ComBase))]
    [DebuggerDisplay("{" + nameof(MemberDeclaration) + "}")]
    public class ComMember : ComBase
    {
        [DataMember(IsRequired = true)]
        public bool IsHidden { get; private set; }

        [DataMember(IsRequired = true)]
        public bool IsRestricted { get; private set; }

        [DataMember(IsRequired = true)]
        public bool ReturnsWithEventsObject { get; private set; }

        [DataMember(IsRequired = true)]
        public bool IsDefault { get; private set; }

        [DataMember(IsRequired = true)]
        public bool IsEnumerator { get; private set; }

        //This member is called on an interface when a bracketed expression is dereferenced.
        [DataMember(IsRequired = true)]
        public bool IsEvaluateFunction { get; private set; }

        [DataMember(IsRequired = true)]
        public ComParameter AsTypeName { get; private set; } = ComParameter.Void;

        [DataMember(IsRequired = true)]
        private List<ComParameter> _parameters = new List<ComParameter>();

        //See https://docs.microsoft.com/en-us/windows/desktop/midl/retval
        //"Parameters with the [retval] attribute are not displayed in user-oriented browsers."
        public IEnumerable<ComParameter> Parameters => _parameters.Where(param => !param.IsReturnValue);

        public ComMember(IComBase parent, ITypeInfo info, FUNCDESC funcDesc) : base(parent, info, funcDesc)
        {
            LoadParameters(funcDesc, info);
            var flags = (FUNCFLAGS)funcDesc.wFuncFlags;
            IsHidden = flags.HasFlag(FUNCFLAGS.FUNCFLAG_FHIDDEN);
            IsRestricted = flags.HasFlag(FUNCFLAGS.FUNCFLAG_FRESTRICTED);
            ReturnsWithEventsObject = flags.HasFlag(FUNCFLAGS.FUNCFLAG_FSOURCE);
            IsDefault = Index == (int)DispId.Value;
            IsEnumerator = Index == (int)DispId.NewEnum;
            IsEvaluateFunction = Index == (int)DispId.Evaluate;
            SetDeclarationType(funcDesc, info);
        }

        private void SetDeclarationType(FUNCDESC funcDesc, ITypeInfo info)
        {
            var returnsHResult = (VarEnum)funcDesc.elemdescFunc.tdesc.vt == VarEnum.VT_HRESULT;
            var returnsVoid = (VarEnum)funcDesc.elemdescFunc.tdesc.vt == VarEnum.VT_VOID;

            if (funcDesc.invkind.HasFlag(INVOKEKIND.INVOKE_PROPERTYGET))
            {
                Type = DeclarationType.PropertyGet;
            }
            else if (funcDesc.invkind.HasFlag(INVOKEKIND.INVOKE_PROPERTYPUT))
            {
                Type = DeclarationType.PropertyLet;
            }
            else if (funcDesc.invkind.HasFlag(INVOKEKIND.INVOKE_PROPERTYPUTREF))
            {
                Type = DeclarationType.PropertySet;
            }
            else if (returnsVoid || !_parameters.Any(param => param.IsReturnValue) && returnsHResult)
            {
                Type = DeclarationType.Procedure;
            }
            else
            {
                Type = DeclarationType.Function;
            }

            if (Type == DeclarationType.Function || Type == DeclarationType.PropertyGet)
            {
                var returnType = new ComParameter(this, funcDesc.elemdescFunc, info, string.Empty);
                if (!_parameters.Any())
                {
                    AsTypeName = returnType;
                }
                else
                {
                    var retval = _parameters.FirstOrDefault(x => x.IsReturnValue);
                    AsTypeName = retval ?? returnType;
                }
            }
        }

        private void LoadParameters(FUNCDESC funcDesc, ITypeInfo info)
        {
            var names = new string[255];
            info.GetNames(Index, names, names.Length, out _);

            for (var index = 0; index < funcDesc.cParams; index++)
            {
                var paramPtr = new IntPtr(funcDesc.lprgelemdescParam.ToInt64() + Marshal.SizeOf(typeof(ELEMDESC)) * index);
                var elemDesc = Marshal.PtrToStructure<ELEMDESC>(paramPtr);
                var param = new ComParameter(this, elemDesc, info, names[index + 1] ?? $"{index}unnamedParameter");
                _parameters.Add(param);
            }

            // See https://docs.microsoft.com/en-us/windows/desktop/midl/propput
            // "A function that has the [propput] attribute must also have, as its last parameter, a parameter that has the [in] attribute."
            if (funcDesc.invkind.HasFlag(INVOKEKIND.INVOKE_PROPERTYPUTREF) ||
                funcDesc.invkind.HasFlag(INVOKEKIND.INVOKE_PROPERTYPUT))
            {
                AsTypeName = _parameters.Last();
                _parameters = _parameters.Take(funcDesc.cParams - 1).ToList();
                return;
            }

            if (Parameters.Any() && funcDesc.cParamsOpt == -1)
            {
                Parameters.Last().IsParamArray = true;
            }
        }

        private string MemberDeclaration
        {
            get
            {
                var type = string.Empty;
                switch (Type)
                {
                    case DeclarationType.Function:
                        type = "Function";
                        break;
                    case DeclarationType.Procedure:
                        type = "Sub";
                        break;
                    case DeclarationType.PropertyGet:
                        type = "Property Get";
                        break;
                    case DeclarationType.PropertyLet:
                        type = "Property Let";
                        break;
                    case DeclarationType.PropertySet:
                        type = "Property Set";
                        break;
                    case DeclarationType.Event:
                        type = "Event";
                        break;
                }
                return $"{(IsHidden || IsRestricted ? "Private" : "Public")} {type} {Name}{(AsTypeName == null ? string.Empty : $" As {AsTypeName.TypeName}")}";
            }
        }
    }
}
