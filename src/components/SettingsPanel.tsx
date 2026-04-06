import React, { useState, useCallback, useEffect } from 'react';
import { cn } from '@/lib/utils';

type AIProvider = 'openai' | 'anthropic' | 'google' | 'openrouter';
type SettingsTab = 'provider' | 'security' | 'about';

interface Settings {
  provider: AIProvider;
  model: string;
  openaiKey: string;
  anthropicKey: string;
  googleKey: string;
  openrouterKey: string;
  openrouterEndpoint: string;
  clearChatOnProviderChange: boolean;
  storeKeysLocally: boolean;
}

interface SettingsPanelProps {
  settings: Settings;
  onSave: (settings: Settings) => void;
  onClose: () => void;
}

const modelOptions: Record<AIProvider, { value: string; label: string; tier?: string }[]> = {
  openai: [
    { value: 'gpt-4o', label: 'GPT-4o', tier: 'Recommended' },
    { value: 'gpt-4o-mini', label: 'GPT-4o Mini', tier: 'Fast' },
    { value: 'gpt-4-turbo', label: 'GPT-4 Turbo' },
    { value: 'gpt-3.5-turbo', label: 'GPT-3.5 Turbo', tier: 'Cheap' },
  ],
  anthropic: [
    { value: 'claude-sonnet-4-20250514', label: 'Claude Sonnet 4', tier: 'Recommended' },
    { value: 'claude-opus-4-20250514', label: 'Claude Opus 4', tier: 'Best' },
    { value: 'claude-3-5-sonnet-20241022', label: 'Claude 3.5 Sonnet' },
    { value: 'claude-3-5-haiku-20241022', label: 'Claude 3.5 Haiku', tier: 'Fast' },
  ],
  google: [
    { value: 'gemini-2.0-flash', label: 'Gemini 2.0 Flash', tier: 'Recommended' },
    { value: 'gemini-2.0-flash-lite', label: 'Gemini 2.0 Flash Lite', tier: 'Fast' },
    { value: 'gemini-1.5-pro', label: 'Gemini 1.5 Pro', tier: 'Best' },
  ],
  openrouter: [
    { value: 'google/gemini-2.0-flash-exp:free', label: 'Gemini 2.0 Flash (Free)', tier: 'Free' },
    { value: 'google/gemini-2.0-flash-lite:free', label: 'Gemini 2.0 Flash Lite (Free)', tier: 'Free' },
    { value: 'meta-llama/llama-3.3-70b-instruct:free', label: 'Llama 3.3 70B (Free)', tier: 'Free' },
    { value: 'qwen/qwen2.5-coder-32b-instruct:free', label: 'Qwen 2.5 Coder 32B (Free)', tier: 'Free' },
    { value: 'mistralai/mistral-7b-instruct:free', label: 'Mistral 7B (Free)', tier: 'Free' },
    { value: 'deepseek/deepseek-chat:free', label: 'DeepSeek Chat (Free)', tier: 'Free' },
    { value: 'openai/gpt-4o', label: 'GPT-4o (via OpenRouter)', tier: 'Recommended' },
    { value: 'anthropic/claude-sonnet-4', label: 'Claude Sonnet 4 (via OpenRouter)', tier: 'Best' },
    { value: 'anthropic/claude-3.5-sonnet', label: 'Claude 3.5 Sonnet (via OpenRouter)' },
    { value: 'openai/gpt-4o-mini', label: 'GPT-4o Mini (via OpenRouter)', tier: 'Cheap' },
    { value: 'meta-llama/llama-3.3-70b-instruct', label: 'Llama 3.3 70B (via OpenRouter)' },
    { value: 'mistralai/mistral-large-2411', label: 'Mistral Large (via OpenRouter)' },
  ],
};

const providerConfig: Record<AIProvider, { name: string; icon: React.ReactNode; color: string; description: string; keyPrefix: string; endpoint?: string }> = {
  openai: {
    name: 'OpenAI',
    icon: (
      <svg className="w-5 h-5" viewBox="0 0 24 24" fill="currentColor">
        <path d="M22.2819 9.8211a5.9847 5.9847 0 0 0-.5157-4.9108 6.0462 6.0462 0 0 0-6.5098-2.9A6.0651 6.0651 0 0 0 4.9807 4.1818a5.9847 5.9847 0 0 0-3.9977 2.9 6.0462 6.0462 0 0 0 .7427 7.0966 5.98 5.98 0 0 0 .511 4.9107 6.051 6.051 0 0 0 6.5145 2.9001A5.9847 5.9847 0 0 0 13.2599 24a6.0557 6.0557 0 0 0 5.7718-4.2058 5.99 5.99 0 0 0 3.9977-2.9001 6.0557 6.0557 0 0 0-.7475-7.0729zm-9.022 12.6081a4.4755 4.4755 0 0 1-2.8764-1.0408l.1397-.0765 5.2525-3.0313a.7948.7948 0 0 0 .3927-.6813v-6.7369l2.0292 1.1718a.198.198 0 0 1 .1042.1491v4.5003a4.4755 4.4755 0 0 1-5.0419 4.4446z" />
      </svg>
    ),
    color: 'from-gray-700 to-gray-900',
    description: 'Direct OpenAI API',
    keyPrefix: 'sk-',
  },
  anthropic: {
    name: 'Anthropic',
    icon: (
      <svg className="w-5 h-5" viewBox="0 0 24 24" fill="currentColor">
        <path d="M17.3599 2.7H21.0039L12.6639 21.3H9.00391L17.3599 2.7ZM2.98391 2.7H6.62791L14.9679 21.3H11.3239L2.98391 2.7Z" />
      </svg>
    ),
    color: 'from-orange-500 to-red-600',
    description: 'Direct Anthropic API',
    keyPrefix: 'sk-ant-',
  },
  google: {
    name: 'Google',
    icon: (
      <svg className="w-5 h-5" viewBox="0 0 24 24" fill="currentColor">
        <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 17.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.39z" />
      </svg>
    ),
    color: 'from-blue-500 to-green-500',
    description: 'Direct Google AI API',
    keyPrefix: 'AIza',
  },
  openrouter: {
    name: 'OpenRouter',
    icon: (
      <svg className="w-5 h-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" />
        <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" />
      </svg>
    ),
    color: 'from-violet-500 to-purple-600',
    description: 'Access 100+ models via one API',
    keyPrefix: 'sk-or-',
    endpoint: 'https://openrouter.ai/api/v1/chat/completions',
  },
};

const defaultSettings: Settings = {
  provider: 'openai',
  model: 'gpt-4o',
  openaiKey: '',
  anthropicKey: '',
  googleKey: '',
  openrouterKey: '',
  openrouterEndpoint: 'https://openrouter.ai/api/v1/chat/completions',
  clearChatOnProviderChange: true,
  storeKeysLocally: true,
};

export const SettingsPanel: React.FC<SettingsPanelProps> = ({ settings, onSave, onClose }) => {
  const [local, setLocal] = useState<Settings>(() => ({ ...defaultSettings, ...settings }));
  const [activeTab, setActiveTab] = useState<SettingsTab>('provider');
  const [showKey, setShowKey] = useState<Record<string, boolean>>({});

  const handleProviderChange = useCallback((provider: AIProvider) => {
    setLocal((prev) => ({
      ...prev,
      provider,
      model: modelOptions[provider][0].value,
    }));
  }, []);

  const handleSave = useCallback(() => {
    onSave(local);
    onClose();
  }, [local, onSave, onClose]);

  const handleClearAllKeys = useCallback(() => {
    setLocal((prev) => ({
      ...prev,
      openaiKey: '',
      anthropicKey: '',
      googleKey: '',
      openrouterKey: '',
    }));
  }, []);

  const currentProvider = providerConfig[local.provider];
  const currentModel = modelOptions[local.provider];

  const tabs: { id: SettingsTab; label: string; icon: React.ReactNode }[] = [
    {
      id: 'provider',
      label: 'AI Provider',
      icon: (
        <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <path d="M12 2L2 7l10 5 10-5-10-5z" />
          <path d="M2 17l10 5 10-5" />
          <path d="M2 12l10 5 10-5" />
        </svg>
      ),
    },
    {
      id: 'security',
      label: 'Security',
      icon: (
        <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
          <path d="M7 11V7a5 5 0 0 1 10 0v4" />
        </svg>
      ),
    },
    {
      id: 'about',
      label: 'About',
      icon: (
        <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <circle cx="12" cy="12" r="10" />
          <path d="M12 16v-4" />
          <path d="M12 8h.01" />
        </svg>
      ),
    },
  ];

  return (
    <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-50 flex items-center justify-center p-4 animate-fade-in">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden animate-scale-in" style={{ maxHeight: '85vh' }}>
        {/* Header */}
        <div className="flex items-center justify-between px-5 py-3.5 border-b border-gray-100">
          <div className="flex items-center gap-2.5">
            <div className="w-7 h-7 rounded-lg bg-gradient-to-br from-[#217346] to-[#185C37] flex items-center justify-center">
              <svg className="w-4 h-4 text-white" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="3" />
                <path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42" />
              </svg>
            </div>
            <h2 className="text-base font-bold text-gray-900">Settings</h2>
          </div>
          <button
            onClick={onClose}
            className="w-7 h-7 rounded-lg hover:bg-gray-100 flex items-center justify-center text-gray-400 hover:text-gray-600 transition-colors"
          >
            <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M18 6L6 18M6 6l12 12" />
            </svg>
          </button>
        </div>

        {/* Tabs */}
        <div className="flex border-b border-gray-100 px-5">
          {tabs.map((tab) => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              className={cn(
                'flex items-center gap-1.5 px-3 py-2.5 text-xs font-medium border-b-2 transition-all',
                activeTab === tab.id
                  ? 'border-[#217346] text-[#217346]'
                  : 'border-transparent text-gray-400 hover:text-gray-600'
              )}
            >
              {tab.icon}
              {tab.label}
            </button>
          ))}
        </div>

        {/* Content */}
        <div className="p-5 overflow-y-auto" style={{ maxHeight: 'calc(85vh - 200px)' }}>
          {activeTab === 'provider' && (
            <div className="space-y-5">
              {/* Provider Selection */}
              <div>
                <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">AI Provider</label>
                <div className="grid grid-cols-2 gap-2">
                  {(Object.keys(providerConfig) as AIProvider[]).map((p) => {
                    const config = providerConfig[p];
                    const isActive = local.provider === p;
                    return (
                      <button
                        key={p}
                        onClick={() => handleProviderChange(p)}
                        className={cn(
                          'flex items-center gap-2.5 px-3 py-2.5 rounded-xl border-2 transition-all text-left',
                          isActive
                            ? 'border-[#217346] bg-[#217346]/5'
                            : 'border-gray-200 hover:border-gray-300'
                        )}
                      >
                        <div className={cn('w-8 h-8 rounded-lg bg-gradient-to-br flex items-center justify-center text-white flex-shrink-0', config.color)}>
                          {config.icon}
                        </div>
                        <div className="min-w-0">
                          <div className={cn('text-xs font-bold', isActive ? 'text-[#217346]' : 'text-gray-700')}>{config.name}</div>
                          <div className="text-[10px] text-gray-400 truncate">{config.description}</div>
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Model Selection */}
              <div>
                <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Model</label>
                <div className="grid grid-cols-1 gap-1.5 max-h-40 overflow-y-auto pr-1">
                  {currentModel.map((m) => (
                    <button
                      key={m.value}
                      onClick={() => setLocal((prev) => ({ ...prev, model: m.value }))}
                      className={cn(
                        'flex items-center justify-between px-3 py-2 rounded-lg border transition-all text-left',
                        local.model === m.value
                          ? 'border-[#217346] bg-[#217346]/5'
                          : 'border-gray-100 hover:border-gray-200 hover:bg-gray-50'
                      )}
                    >
                      <div>
                        <div className={cn('text-xs font-medium', local.model === m.value ? 'text-[#217346]' : 'text-gray-700')}>{m.label}</div>
                      </div>
                      {m.tier && (
                        <span className={cn(
                          'text-[9px] font-semibold px-1.5 py-0.5 rounded-full',
                          m.tier === 'Recommended' ? 'bg-emerald-100 text-emerald-700' :
                          m.tier === 'Best' ? 'bg-purple-100 text-purple-700' :
                          m.tier === 'Fast' ? 'bg-blue-100 text-blue-700' :
                          m.tier === 'Free' ? 'bg-green-100 text-green-700' :
                          m.tier === 'Cheap' ? 'bg-amber-100 text-amber-700' :
                          'bg-gray-100 text-gray-600'
                        )}>
                          {m.tier}
                        </span>
                      )}
                    </button>
                  ))}
                </div>
              </div>

              {/* API Key for selected provider */}
              <div>
                <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">API Key</label>
                <div className="relative">
                  <input
                    type={showKey[local.provider] ? 'text' : 'password'}
                    value={local.provider === 'openai' ? local.openaiKey : local.provider === 'anthropic' ? local.anthropicKey : local.provider === 'google' ? local.googleKey : local.openrouterKey}
                    onChange={(e) => {
                      const key = local.provider === 'openai' ? 'openaiKey' : local.provider === 'anthropic' ? 'anthropicKey' : local.provider === 'google' ? 'googleKey' : 'openrouterKey';
                      setLocal((prev) => ({ ...prev, [key]: e.target.value }));
                    }}
                    placeholder={`${currentProvider.keyPrefix}...`}
                    className="w-full px-3 py-2.5 pr-10 rounded-xl border border-gray-200 bg-gray-50 text-xs font-mono text-gray-800 placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-[#217346]/20 focus:border-[#217346] transition-all"
                  />
                  <button
                    onClick={() => setShowKey((prev) => ({ ...prev, [local.provider]: !prev[local.provider] }))}
                    className="absolute right-2 top-1/2 -translate-y-1/2 w-6 h-6 rounded flex items-center justify-center text-gray-400 hover:text-gray-600 transition-colors"
                  >
                    {showKey[local.provider] ? (
                      <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24" />
                        <line x1="1" y1="1" x2="23" y2="23" />
                      </svg>
                    ) : (
                      <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
                        <circle cx="12" cy="12" r="3" />
                      </svg>
                    )}
                  </button>
                </div>
                <p className="text-[10px] text-gray-400 mt-1">Your key is stored locally and never sent to any third party.</p>
              </div>

              {/* OpenRouter Endpoint */}
              {local.provider === 'openrouter' && (
                <div>
                  <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Endpoint URL</label>
                  <input
                    type="text"
                    value={local.openrouterEndpoint}
                    onChange={(e) => setLocal((prev) => ({ ...prev, openrouterEndpoint: e.target.value }))}
                    placeholder="https://openrouter.ai/api/v1/chat/completions"
                    className="w-full px-3 py-2.5 rounded-xl border border-gray-200 bg-gray-50 text-xs font-mono text-gray-800 placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-[#217346]/20 focus:border-[#217346] transition-all"
                  />
                  <p className="text-[10px] text-gray-400 mt-1">Default: https://openrouter.ai/api/v1/chat/completions</p>
                </div>
              )}
            </div>
          )}

          {activeTab === 'security' && (
            <div className="space-y-5">
              {/* Key Management */}
              <div className="p-3 rounded-xl bg-amber-50 border border-amber-100">
                <div className="flex items-start gap-2">
                  <svg className="w-4 h-4 text-amber-500 mt-0.5 flex-shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
                    <line x1="12" y1="9" x2="12" y2="13" />
                    <line x1="12" y1="17" x2="12.01" y2="17" />
                  </svg>
                  <div>
                    <h4 className="text-xs font-semibold text-amber-800">API Key Security</h4>
                    <p className="text-[10px] text-amber-600 mt-0.5 leading-relaxed">
                      Your API keys are stored in your browser's localStorage. They are never transmitted to any server except the AI provider's API endpoint. We do not collect, store, or share your keys.
                    </p>
                  </div>
                </div>
              </div>

              {/* Data Privacy */}
              <div>
                <h4 className="text-xs font-semibold text-gray-700 mb-2">Data Privacy</h4>
                <div className="space-y-2">
                  <div className="flex items-center justify-between p-2.5 rounded-lg bg-gray-50 border border-gray-100">
                    <div>
                      <div className="text-xs font-medium text-gray-700">Local Storage Only</div>
                      <div className="text-[10px] text-gray-400">All data stays in your browser</div>
                    </div>
                    <div className="w-8 h-4 rounded-full bg-emerald-500 relative">
                      <div className="absolute right-0.5 top-0.5 w-3 h-3 rounded-full bg-white shadow-sm" />
                    </div>
                  </div>
                  <div className="flex items-center justify-between p-2.5 rounded-lg bg-gray-50 border border-gray-100">
                    <div>
                      <div className="text-xs font-medium text-gray-700">No Telemetry</div>
                      <div className="text-[10px] text-gray-400">No usage data is collected</div>
                    </div>
                    <div className="w-8 h-4 rounded-full bg-emerald-500 relative">
                      <div className="absolute right-0.5 top-0.5 w-3 h-3 rounded-full bg-white shadow-sm" />
                    </div>
                  </div>
                  <div className="flex items-center justify-between p-2.5 rounded-lg bg-gray-50 border border-gray-100">
                    <div>
                      <div className="text-xs font-medium text-gray-700">No Cloud Sync</div>
                      <div className="text-[10px] text-gray-400">Settings never leave your device</div>
                    </div>
                    <div className="w-8 h-4 rounded-full bg-emerald-500 relative">
                      <div className="absolute right-0.5 top-0.5 w-3 h-3 rounded-full bg-white shadow-sm" />
                    </div>
                  </div>
                </div>
              </div>

              {/* Actions */}
              <div>
                <h4 className="text-xs font-semibold text-gray-700 mb-2">Actions</h4>
                <div className="space-y-2">
                  <button
                    onClick={handleClearAllKeys}
                    className="w-full px-3 py-2.5 rounded-xl border border-red-200 bg-red-50 text-xs font-medium text-red-600 hover:bg-red-100 transition-colors flex items-center justify-center gap-2"
                  >
                    <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M3 6h18M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6" />
                    </svg>
                    Clear All API Keys
                  </button>
                  <button
                    onClick={() => {
                      localStorage.clear();
                      setLocal({ ...defaultSettings });
                    }}
                    className="w-full px-3 py-2.5 rounded-xl border border-red-200 bg-red-50 text-xs font-medium text-red-600 hover:bg-red-100 transition-colors flex items-center justify-center gap-2"
                  >
                    <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M3 6h18M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6" />
                    </svg>
                    Clear All Data (localStorage)
                  </button>
                </div>
              </div>

              {/* Connection Info */}
              <div>
                <h4 className="text-xs font-semibold text-gray-700 mb-2">Connection Details</h4>
                <div className="p-3 rounded-lg bg-gray-50 border border-gray-100 space-y-1.5">
                  <div className="flex justify-between">
                    <span className="text-[10px] text-gray-400">Provider</span>
                    <span className="text-[10px] font-mono text-gray-600">{currentProvider.name}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-[10px] text-gray-400">Model</span>
                    <span className="text-[10px] font-mono text-gray-600">{local.model}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-[10px] text-gray-400">Key Status</span>
                    <span className={cn('text-[10px] font-mono', (local.provider === 'openai' ? local.openaiKey : local.provider === 'anthropic' ? local.anthropicKey : local.provider === 'google' ? local.googleKey : local.openrouterKey) ? 'text-emerald-600' : 'text-red-500')}>
                      {(local.provider === 'openai' ? local.openaiKey : local.provider === 'anthropic' ? local.anthropicKey : local.provider === 'google' ? local.googleKey : local.openrouterKey) ? 'Configured' : 'Not set'}
                    </span>
                  </div>
                  {local.provider === 'openrouter' && (
                    <div className="flex justify-between">
                      <span className="text-[10px] text-gray-400">Endpoint</span>
                      <span className="text-[10px] font-mono text-gray-600 truncate ml-2 max-w-[200px]">{local.openrouterEndpoint}</span>
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {activeTab === 'about' && (
            <div className="space-y-5">
              <div className="text-center py-4">
                <div className="w-12 h-12 rounded-xl bg-gradient-to-br from-[#217346] to-[#185C37] flex items-center justify-center mx-auto mb-3 shadow-lg shadow-[#217346]/20">
                  <svg className="w-6 h-6 text-white" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M12 2L2 7l10 5 10-5-10-5z" />
                    <path d="M2 17l10 5 10-5" />
                    <path d="M2 12l10 5 10-5" />
                  </svg>
                </div>
                <h3 className="text-base font-bold text-gray-900">Cel</h3>
                <p className="text-xs text-gray-400 mt-0.5">Agentic AI for Excel</p>
                <p className="text-[10px] text-gray-400 mt-1">Version 1.0.0</p>
              </div>

              <div className="space-y-3">
                <div className="p-3 rounded-xl bg-gray-50 border border-gray-100">
                  <h4 className="text-xs font-semibold text-gray-700 mb-1">What is Cel?</h4>
                  <p className="text-[11px] text-gray-500 leading-relaxed">
                    An agentic AI Excel add-in that lets you chat with intelligent agents to read, edit, format, and analyze your spreadsheets directly from a sidebar.
                  </p>
                </div>

                <div className="p-3 rounded-xl bg-gray-50 border border-gray-100">
                  <h4 className="text-xs font-semibold text-gray-700 mb-1">Supported AI Providers</h4>
                  <div className="space-y-1 mt-2">
                    <div className="flex items-center gap-2 text-[11px] text-gray-500">
                      <div className="w-1.5 h-1.5 rounded-full bg-gray-400" />
                      OpenAI (GPT-4o, GPT-4o Mini, GPT-4 Turbo)
                    </div>
                    <div className="flex items-center gap-2 text-[11px] text-gray-500">
                      <div className="w-1.5 h-1.5 rounded-full bg-gray-400" />
                      Anthropic (Claude Sonnet 4, Opus 4, 3.5 Sonnet)
                    </div>
                    <div className="flex items-center gap-2 text-[11px] text-gray-500">
                      <div className="w-1.5 h-1.5 rounded-full bg-gray-400" />
                      Google (Gemini 2.0 Flash, 1.5 Pro)
                    </div>
                    <div className="flex items-center gap-2 text-[11px] text-gray-500">
                      <div className="w-1.5 h-1.5 rounded-full bg-gray-400" />
                      OpenRouter (100+ models including Llama, Mistral, DeepSeek)
                    </div>
                  </div>
                </div>

                <div className="p-3 rounded-xl bg-gray-50 border border-gray-100">
                  <h4 className="text-xs font-semibold text-gray-700 mb-1">Capabilities</h4>
                  <div className="space-y-1 mt-2">
                    <div className="flex items-center gap-2 text-[11px] text-gray-500">
                      <div className="w-1.5 h-1.5 rounded-full bg-[#217346]" />
                      Read & analyze data from any range or sheet
                    </div>
                    <div className="flex items-center gap-2 text-[11px] text-gray-500">
                      <div className="w-1.5 h-1.5 rounded-full bg-[#217346]" />
                      Write values and formulas to cells
                    </div>
                    <div className="flex items-center gap-2 text-[11px] text-gray-500">
                      <div className="w-1.5 h-1.5 rounded-full bg-[#217346]" />
                      Apply formatting (bold, colors, alignment)
                    </div>
                    <div className="flex items-center gap-2 text-[11px] text-gray-500">
                      <div className="w-1.5 h-1.5 rounded-full bg-[#217346]" />
                      Create charts (column, bar, line, pie, etc.)
                    </div>
                    <div className="flex items-center gap-2 text-[11px] text-gray-500">
                      <div className="w-1.5 h-1.5 rounded-full bg-[#217346]" />
                      Add/delete sheets, rows, columns
                    </div>
                    <div className="flex items-center gap-2 text-[11px] text-gray-500">
                      <div className="w-1.5 h-1.5 rounded-full bg-[#217346]" />
                      Create tables, sort data, autofill formulas
                    </div>
                  </div>
                </div>

                <div className="p-3 rounded-xl bg-gray-50 border border-gray-100">
                  <h4 className="text-xs font-semibold text-gray-700 mb-1">Architecture</h4>
                  <p className="text-[11px] text-gray-500 leading-relaxed">
                    Uses CrewAI multi-agent orchestration (Planner → Executor → Validator) running on a local FastAPI backend. Excel operations are executed via Office.js in the browser for security and direct workbook access.
                  </p>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-5 py-3 border-t border-gray-100 bg-gray-50/50 flex gap-2">
          <button
            onClick={onClose}
            className="flex-1 px-3 py-2 rounded-xl border border-gray-200 text-xs font-medium text-gray-600 hover:bg-gray-100 transition-colors"
          >
            Cancel
          </button>
          <button
            onClick={handleSave}
            className="flex-1 px-3 py-2 rounded-xl bg-gradient-to-r from-[#217346] to-[#185C37] text-xs font-medium text-white hover:from-[#1E6B41] hover:to-[#145030] transition-all shadow-sm shadow-[#217346]/20"
          >
            Save Settings
          </button>
        </div>
      </div>
    </div>
  );
};
