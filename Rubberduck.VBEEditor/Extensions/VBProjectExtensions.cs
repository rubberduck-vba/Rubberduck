using IOException = System.IO.IOException;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Extensions
{
    public static class VBProjectExtensions
    {
        public static bool TryGetFullPath(this IVBProject project, out string fullPath)
        {
            try
            {
                fullPath = project.FileName;
            }
            catch (IOException)
            {
                // Filename throws exception if unsaved.
                fullPath = null;
                return false;
            }
            catch (COMException e)
            {
                LogManager.GetLogger(typeof(IVBProject).FullName).Warn(e);
                fullPath = null;
                return false;
            }

            return true;
        }
    }
}