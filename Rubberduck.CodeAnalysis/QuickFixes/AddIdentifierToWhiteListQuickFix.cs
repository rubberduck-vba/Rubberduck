using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class AddIdentifierToWhiteListQuickFix : QuickFixBase
    {
        private readonly IPersistanceService<CodeInspectionSettings> _settings;

        public AddIdentifierToWhiteListQuickFix(IPersistanceService<CodeInspectionSettings> settings)
            : base(typeof(HungarianNotationInspection), typeof(UseMeaningfulNameInspection))
        {
            _settings = settings;
        }

        //The rewriteSession is optional since it is not used in this particular quickfix.
        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession = null)
        {
            var inspectionSettings = _settings.Load(new CodeInspectionSettings()) ?? new CodeInspectionSettings();
            var whitelist = inspectionSettings.WhitelistedIdentifiers;
            inspectionSettings.WhitelistedIdentifiers =
                whitelist.Concat(new[] { new WhitelistedIdentifierSetting(result.Target.IdentifierName) }).ToArray();
            _settings.Save(inspectionSettings);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.WhiteListIdentifierQuickFix;

        public override bool CanFixInProcedure { get; } = false;
        public override bool CanFixInModule { get; } = false;
        public override bool CanFixInProject { get; } = false;
    }
}
