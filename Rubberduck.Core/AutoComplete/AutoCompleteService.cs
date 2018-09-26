using System;
using System.Collections.Generic;
using System.Diagnostics;
using Rubberduck.AutoComplete.SelfClosingPairCompletion;
using Rubberduck.Settings;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteService : IDisposable
    {
        private readonly IGeneralConfigService _configService;
        private readonly List<SelfClosingPair> _selfClosingPairs = new List<SelfClosingPair>
        {
            new SelfClosingPair('(', ')'),
            new SelfClosingPair('"', '"'),
            new SelfClosingPair('[', ']'),
            new SelfClosingPair('{', '}'),
        };

        private readonly AutoCompleteKeyDownHandler _handler;
        private readonly SelfClosingPairCompletionService _selfClosingPairCompletion;
        private readonly ICodeStringPrettifier _prettifier;

        private AutoCompleteSettings _settings;
        private bool _popupShown;
        private bool _enabled;
        private bool _initialized;

        public AutoCompleteService(IGeneralConfigService configService, SelfClosingPairCompletionService selfClosingPairCompletion, ICodeStringPrettifier prettifier)
        {
            _handler = new AutoCompleteKeyDownHandler(()=>_settings, ()=>_selfClosingPairs, ()=>_selfClosingPairCompletion, ()=>_prettifier);

            _selfClosingPairCompletion = selfClosingPairCompletion;
            _prettifier = prettifier;
            _configService = configService;
            _configService.SettingsChanged += ConfigServiceSettingsChanged;
        }

        public void Enable()
        {
            if (!_initializing)
            {
                InitializeConfig();
            }

            if (!_enabled)
            {
                VBENativeServices.KeyDown += HandleKeyDown;
                VBENativeServices.IntelliSenseChanged += HandleIntelliSenseChanged;
                _enabled = true;
            }
        }

        private bool _initializing;
        private void InitializeConfig()
        {
            _initializing = true;
            // No reason to think this would throw, but if it does, _initializing state needs to be reset.
            try
            {
                if (!_initialized)
                {
                    var config = _configService.LoadConfiguration();
                    ApplyAutoCompleteSettings(config);
                }
            }
            finally
            {
                _initializing = false;
            }            
        }

        public void Disable()
        {
            if (_enabled && _initialized)
            {
                VBENativeServices.KeyDown -= HandleKeyDown;
                VBENativeServices.IntelliSenseChanged -= HandleIntelliSenseChanged;
                _enabled = false;
            }
        }

        private void HandleIntelliSenseChanged(object sender, IntelliSenseEventArgs e)
        {
            _popupShown = e.Visible;
        }

        private void ConfigServiceSettingsChanged(object sender, ConfigurationChangedEventArgs e)
        {
            var config = _configService.LoadConfiguration();
            ApplyAutoCompleteSettings(config);
        }
        
        public void ApplyAutoCompleteSettings(Configuration config)
        {
            _settings = config.UserSettings.AutoCompleteSettings;
            if (_settings.IsEnabled)
            {
                Enable();
            }
            else
            {
                Disable();
            }
            _initialized = true;
        }

        private void HandleKeyDown(object sender, AutoCompleteEventArgs e)
        {
            if (e.Character == default && !e.IsDelete)
            {
                return;
            }

            var module = e.CodeModule;
            var qualifiedSelection = module.GetQualifiedSelection();
            Debug.Assert(qualifiedSelection != null, nameof(qualifiedSelection) + " != null");
            var pSelection = qualifiedSelection.Value.Selection;

            if (_popupShown || pSelection.LineCount > 1 || e.IsDelete)
            {
                return;
            }

            _handler.Run(module, pSelection, e);
        }

        public void Dispose()
        {
            Disable();
            if (_configService != null)
            {
                _configService.SettingsChanged -= ConfigServiceSettingsChanged;
            }
        }
    }
}
