using Rubberduck.Refactorings.MoveMember;

namespace Rubberduck.Refactorings
{
    public interface IMovedContentProviderFactory
    {
        IMovedContentProvider CreateDefaultProvider();
        IMovedContentProvider CreatePreviewProvider();
    }

    public class MovedContentProviderFactory : IMovedContentProviderFactory
    {
        public IMovedContentProvider CreateDefaultProvider()
        {
            return new MovedContentProvider();
        }
        public IMovedContentProvider CreatePreviewProvider()
        {
            return new MovedContentPreviewProvider();
        }
    }
}
