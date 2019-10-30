using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.JunkDrawer.Hacks
{
    public static class VarDescExtensions
    {
        /// <remarks>
        /// Use only with VBA-supplied <see cref="ITypeInfo"/> which may return a <see cref="VARDESC"/> that do not conform to 
        /// the MS-OAUT in describing the constants. See section 2.2.43 at: https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/ae7791d2-4399-4dff-b7c6-b0d4f3dce982
        /// 
        /// To further complicate the situation, on 64-bit platform, the <see cref="VARDESC.DESCUNION.oInst"/> is a 32-bit integer whereas
        /// the <see cref="VARDESC.DESCUNION.lpvarValue"/> is a pointer. On 32-bit platform, the sizes of 2 members are exactly same so no
        /// problem. But on 64-bit platform, setting the <c>oInst</c>to 0 does not necessarily zero-initialize the entire region. Thus, the 
        /// upper 32-bit part of the <c>lpvarValue</c> can contain garbage which will confound the simple null pointer check. Thus to guard 
        /// against this, we will check the <c>oInst</c> value to see if it's zero. 
        /// 
        /// There is a small but non-zero chance that there might be a valid pointer that happens to be only in high half of the address...
        /// in that case, it'll be wrong but since VBA is always writing <see cref="VARKIND.VAR_STATIC"/> to the <see cref="VARDESC.varkind"/>
        /// field, we're kind of stuck...
        /// </remarks>
        /// <param name="varDesc">The <see cref="VARDESC"/> from a VBA <see cref="ITypeInfo"/></param>
        /// <returns>True if this is most likely a constant. False when it's definitely not.</returns>
        public static bool IsValidVBAConstant(this VARDESC varDesc)
        {
            return varDesc.varkind == VARKIND.VAR_STATIC && varDesc.desc.oInst != 0;
        }
    }
}
