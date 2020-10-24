
namespace Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences
{
    public struct PrivateUDTMemberReferenceReplacementExpressions
    {
        public PrivateUDTMemberReferenceReplacementExpressions(string memberAccessExpression)
        {
            MemberAccesExpression = memberAccessExpression;
            _localReferenceExpression = memberAccessExpression;
        }

        public string MemberAccesExpression { set; get; }

        private string _localReferenceExpression;
        public string LocalReferenceExpression
        {
            set => _localReferenceExpression = value;
            get => _localReferenceExpression ?? MemberAccesExpression;
        }
    }
}
