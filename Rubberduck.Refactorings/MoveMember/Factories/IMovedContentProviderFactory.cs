using Rubberduck.Refactorings.MoveMember;

namespace Rubberduck.Refactorings
{
    public interface IMovedContentProviderFactory
    {
        INewContentProvider CreateDefaultProvider();
        INewContentProvider CreatePreviewProvider();
    }

    public class MovedContentProviderFactory : IMovedContentProviderFactory
    {
        public INewContentProvider CreateDefaultProvider()
        {
            return new NewContentProvider();
        }
        public INewContentProvider CreatePreviewProvider()
        {
            return new NewContentPreviewProvider();
        }
    }
}
