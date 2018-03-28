using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Extensions
{
    public static class VBComponentExtensions
    {
        public static bool HasEqualCodeModule(this IVBComponent component, IVBComponent otherComponent)
        {
            using (var otherCodeModule = otherComponent.CodeModule)
            {
                return component.HasEqualCodeModule(otherCodeModule);
            }
        }

        public static bool HasEqualCodeModule(this IVBComponent component, ICodeModule otherCodeModule)
        {
            using (var codeModule = component.CodeModule)
            {
                return codeModule.Equals(otherCodeModule);
            }
        }
    }
}
