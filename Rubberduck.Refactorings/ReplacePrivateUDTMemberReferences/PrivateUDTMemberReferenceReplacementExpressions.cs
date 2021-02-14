
namespace Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences
{
    public struct PrivateUDTMemberReferenceReplacementExpressions
    {
        public PrivateUDTMemberReferenceReplacementExpressions(string propertyAccessExpression)
        {
            MemberAccesExpression = propertyAccessExpression;
            _udtMemberLocalReferenceExpression = propertyAccessExpression;
        }

        public string MemberAccesExpression { set; get; }

        private string _udtMemberLocalReferenceExpression;
        public string UDTMemberInternalReferenceExpression
        {
            set => _udtMemberLocalReferenceExpression = value;
            get => _udtMemberLocalReferenceExpression ?? MemberAccesExpression;
        }
    }
}
