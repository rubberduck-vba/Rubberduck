using Org.Benf.OleWoo.Listeners;
using Org.Benf.OleWoo.Typelib;

namespace Rubberduck.Deployment.Build.IdlGeneration
{
    public class IdlListener : TypeLibListenerBase
    {
        public override void EnterTypeLib(ITlibNode libNode)
        {
            libNode.Data.Name = libNode.Data.Name.Replace("_", string.Empty);
            libNode.Data.ShortName = libNode.Data.ShortName.Replace("_", string.Empty);
        }

        public override void EnterCoClass(ITlibNode coClassNode)
        {
            if (coClassNode.ShortName.StartsWith("_"))
            {
                coClassNode.Data.Attributes.Add("hidden");
                coClassNode.Data.Attributes.Add("restricted");
            }
        }

        public override void EnterCoClassInterface(ITlibNode coClassInterfaceNode)
        {
            if (coClassInterfaceNode.Parent.ShortName.StartsWith("_")) //&& coClassInterfaceNode.ShortName != "_Object")
            {
                coClassInterfaceNode.Data.Attributes.Remove("default");
                coClassInterfaceNode.Data.Attributes.Add("restricted");
            }
        }

        public override void EnterInterface(ITlibNode interfaceNode)
        {
            if (interfaceNode.ShortName.StartsWith("_") || interfaceNode.ShortName == "IDockableWindowHost")
            {
                interfaceNode.Data.Attributes.Add("restricted");
            }
        }

        public override void EnterEnumValue(ITlibNode enumValueNode)
        {
            var unwantedPrefix = enumValueNode.Parent.Data.ShortName + "_";
            enumValueNode.Data.Name = enumValueNode.Data.Name.Replace(unwantedPrefix, "rd");
            enumValueNode.Data.ShortName = enumValueNode.Data.ShortName.Replace(unwantedPrefix, "rd");
        }
    }
}
