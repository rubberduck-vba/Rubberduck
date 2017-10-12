using System;
using System.ComponentModel;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Abstract
{
    public interface IControl : ISafeComWrapper, IEquatable<IControl>
    {
        string Name { get; set; }
    }

    public static class ControlExtensions
    {
        public static string TypeName(this IControl control)
        {
            try
            {
                return TypeDescriptor.GetClassName(control.Target);
            }
            catch
            {
                return "Control";
            }
        }     
    }
}