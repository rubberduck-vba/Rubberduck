using Rubberduck.Refactorings.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings
{
    public interface IConflictDetectionSessionDataFactory
    {
        IConflictDetectionSessionData Create();
    }

    public class ConflictDetectionSessionDataFactory : IConflictDetectionSessionDataFactory
    {
        private readonly IConflictDetectionDeclarationProxyFactory _proxyFactory;
        public ConflictDetectionSessionDataFactory(IConflictDetectionDeclarationProxyFactory proxyFactory)
        {
            _proxyFactory = proxyFactory;
        }

        public IConflictDetectionSessionData Create()
        {
            return new ConflictDetectionSessionData(_proxyFactory);
        }
    }

}
