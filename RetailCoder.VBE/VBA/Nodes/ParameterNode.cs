using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Nodes
{
    public class ParameterNode : Node
    {
        public enum VBParameterType
        {
            ImplicitByRef,
            ByRef,
            ByVal
        }

        public ParameterNode(VBParser.ArgContext context, string scope)
            : base(context, scope)
        {
        }

        private new VBParser.ArgContext Context { get { return Context;} }

        public VBParameterType PassedBy
        {
            get
            {
                return Context.BYVAL() != null 
                    ? VBParameterType.ByVal
                    : Context.BYREF() != null
                        ? VBParameterType.ByRef 
                        : VBParameterType.ImplicitByRef;
            }
        }

        public string Name { get { return Context.AmbiguousIdentifier().GetText(); } }
        public string TypeName 
        { 
            get 
            { 
                return Context.AsTypeClause() == null || string.IsNullOrEmpty(Context.AsTypeClause().GetText())
                    ? Tokens.Variant
                    : Context.AsTypeClause().Type().GetText(); 
            } 
        }

        public bool IsOptional
        {
            get
            {
                return Context.OPTIONAL() == null || string.IsNullOrEmpty(Context.OPTIONAL().GetText());
            }
        }
    }
}
