﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using Rubberduck.Settings;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.AutoComplete.Service
{
    public class AutoCompleteService : IDisposable
    {
        private readonly IGeneralConfigService _configService;
        private readonly IEnumerable<AutoCompleteHandlerBase> _handlers;

        private AutoCompleteSettings _settings;
        private bool _popupShown;
        private bool _enabled;
        private bool _initialized;

        public AutoCompleteService(IGeneralConfigService configService, IEnumerable<AutoCompleteHandlerBase> handlers)
        {
            _configService = configService;
            _configService.SettingsChanged += ConfigServiceSettingsChanged;

            _handlers = handlers;
            InitializeConfig(); // I/O work in constructor?
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

        private bool WillHandle(AutoCompleteEventArgs e)
        {
            Debug.Assert(_settings != null);

            if (!_enabled)
            {
                Debug.Assert(_enabled, "KeyDown controller is executing, but auto-completion service is disabled.");
                return false; // release build should bail out here!
            }

            if (_popupShown || e.Character == default && e.IsDeleteKey)
            {
                return false;
            }

            var module = e.Module;
            using (var pane = module.CodePane)
            {
                if (pane.Selection.LineCount > 1)
                {
                    return false;
                }
            }

            return true;
        }

        private void HandleKeyDown(object sender, AutoCompleteEventArgs e)
        {
            if (!WillHandle(e))
            {
                return;
            }

            foreach (var handler in _handlers)
            {
                var result = handler.Handle(e, _settings);
                if (result != null && e.Handled)
                {
                    return;
                }
            }
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
