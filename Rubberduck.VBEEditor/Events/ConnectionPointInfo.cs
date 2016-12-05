using System;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class ConnectionPointInfo
    {
        public ConnectionPointInfo(IConnectionPoint connectionPoint)
        {
            _connectionPoint = connectionPoint;
        }

        private readonly IConnectionPoint _connectionPoint;
        private int? _cookie;

        public bool HasConnectionPoint { get { return _connectionPoint != null; } }

        public void Advise(IVBComponentsEventsSink componentsEventsSink)
        {
            if (_cookie.HasValue) { throw new InvalidOperationException(); }
            int cookie;
            _connectionPoint.Advise(componentsEventsSink, out cookie);
            _cookie = cookie;
        }

        public void Advise(IVBProjectsEventsSink projectsEventsSink)
        {
            if (_cookie.HasValue) { throw new InvalidOperationException(); }
            int cookie;
            _connectionPoint.Advise(projectsEventsSink, out cookie);
            _cookie = cookie;
        }

        public void Unadvise()
        {
            if (!_cookie.HasValue) { throw new InvalidOperationException(); }
            try
            {
                _connectionPoint.Unadvise(_cookie.Value);
            }
            catch (InvalidOperationException)
            {
                // hey, we tried.
            }
        }
    }
}