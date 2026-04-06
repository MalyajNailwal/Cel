import React, { useState, useCallback } from 'react';
import { cn } from '@/lib/utils';
import type { ChatMessage } from '@/lib/ai-providers';
import type { ToolCall, ToolResult } from '@/lib/ai-providers';

interface MessageBubbleProps {
  message: ChatMessage;
  toolCalls?: ToolCall[];
  toolResults?: ToolResult[];
  planSteps?: { action: string; params: Record<string, any>; description: string }[];
  executionResults?: { action: string; success: boolean; output: string }[];
  chartImage?: string;
  chartImages?: string[];
}

export const MessageBubble: React.FC<MessageBubbleProps> = ({
  message,
  toolCalls,
  toolResults,
  planSteps,
  executionResults,
  chartImage,
  chartImages,
}) => {
  const isUser = message.role === 'user';
  const [copied, setCopied] = useState(false);

  const handleCopy = useCallback(() => {
    navigator.clipboard.writeText(message.content);
    setCopied(true);
    setTimeout(() => setCopied(false), 1500);
  }, [message.content]);

  return (
    <div className={cn('flex gap-2 mb-3 animate-slide-up', isUser ? 'flex-row-reverse' : 'flex-row')}>
      <div
        className={cn(
          'flex-shrink-0 w-7 h-7 rounded-full flex items-center justify-center text-[10px] font-bold',
          isUser
            ? 'bg-gradient-to-br from-[#217346] to-[#185C37] text-white shadow-green'
            : 'bg-gradient-to-br from-[#217346]/10 to-[#217346]/5 text-[#217346] ring-1 ring-[#217346]/10'
        )}
      >
        {isUser ? (
          <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2" />
            <circle cx="12" cy="7" r="4" />
          </svg>
        ) : (
          <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M12 2L2 7l10 5 10-5-10-5z" />
            <path d="M2 17l10 5 10-5" />
            <path d="M2 12l10 5 10-5" />
          </svg>
        )}
      </div>
      <div className={cn('flex flex-col max-w-[85%] min-w-0', isUser ? 'items-end' : 'items-start')}>
        {planSteps && planSteps.length > 0 && (
          <AgentWorkflowDisplay steps={planSteps} results={executionResults || []} />
        )}
        {toolCalls && toolCalls.length > 0 && !planSteps && (
          <ToolCallDisplay toolCalls={toolCalls} toolResults={toolResults} />
        )}
        <div className="group relative">
          <div
            className={cn(
              'px-3.5 py-2.5 text-[13px] leading-relaxed break-words overflow-wrap-anywhere whitespace-pre-wrap',
              isUser
                ? 'bg-gradient-to-br from-[#217346] to-[#185C37] text-white rounded-xl rounded-tr-sm shadow-green'
                : 'bg-white text-gray-700 rounded-xl rounded-tl-sm border border-gray-200/80 shadow-soft'
            )}
          >
            {message.content}
          </div>
          {isUser && (
            <button
              onClick={handleCopy}
              className={cn(
                'absolute top-1 right-1 p-1.5 rounded-md transition-all opacity-0 group-hover:opacity-100',
                copied ? 'bg-emerald-500 text-white' : 'bg-white/90 text-gray-500 hover:text-[#217346] hover:bg-white'
              )}
              title={copied ? 'Copied!' : 'Copy'}
            >
              {copied ? (
                <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <polyline points="20 6 9 17 4 12" />
                </svg>
              ) : (
                <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <rect x="9" y="9" width="13" height="13" rx="2" />
                  <path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1" />
                </svg>
              )}
            </button>
          )}
        </div>
        {chartImages && chartImages.length > 1 && (
          <div className="mt-2 space-y-3">
            {chartImages.map((img, idx) => (
              <div key={idx} className="p-2 bg-gray-50 rounded-lg border border-gray-200">
                <img src={img} alt={`Chart ${idx + 1}`} className="max-w-full h-auto rounded" />
              </div>
            ))}
          </div>
        )}
        {chartImage && (!chartImages || chartImages.length <= 1) && (
          <div className="mt-2 p-2 bg-gray-50 rounded-lg border border-gray-200">
            <img src={chartImage} alt="Chart" className="max-w-full h-auto rounded" />
          </div>
        )}
      </div>
    </div>
  );
};

const AgentWorkflowDisplay: React.FC<{
  steps: { action: string; params: Record<string, any>; description: string }[];
  results: { action: string; success: boolean; output: string }[];
}> = ({ steps, results }) => {
  const [expanded, setExpanded] = useState(false);

  const getActionColor = (action: string): string => {
    if (action.startsWith('get_')) return 'from-blue-400 to-blue-500';
    if (action.startsWith('set_') || action === 'auto_fill') return 'from-emerald-400 to-emerald-500';
    if (action === 'apply_format') return 'from-amber-400 to-amber-500';
    if (action === 'create_chart') return 'from-purple-400 to-purple-500';
    return 'from-gray-400 to-gray-500';
  };

  const getActionIcon = (action: string): React.ReactNode => {
    if (action.startsWith('get_')) return '📊';
    if (action.startsWith('set_')) return '✏️';
    if (action === 'apply_format') return '🎨';
    if (action === 'create_chart') return '📈';
    return '⚙️';
  };

  return (
    <div className="mb-2 w-full max-w-full min-w-0 animate-scale-in">
      <button
        onClick={() => setExpanded(!expanded)}
        className="flex items-center gap-2 px-2.5 py-1.5 rounded-lg glass-green border border-[#217346]/10 text-[11px] text-[#217346] hover:bg-[#217346]/10 transition-all w-full"
      >
        <div className="flex items-center -space-x-1">
          {steps.slice(0, 4).map((step, i) => {
            const result = results[i];
            return (
              <div
                key={i}
                className={cn(
                  'w-5 h-5 rounded-full bg-gradient-to-br flex items-center justify-center text-[8px] relative ring-2 ring-white',
                  getActionColor(step.action)
                )}
              >
                {getActionIcon(step.action)}
                {result && (
                  <div className={cn(
                    'absolute -top-0.5 -right-0.5 w-2.5 h-2.5 rounded-full border border-white flex items-center justify-center text-[6px] font-bold',
                    result.success ? 'bg-emerald-500 text-white' : 'bg-red-500 text-white'
                  )}>
                    {result.success ? '✓' : '✗'}
                  </div>
                )}
              </div>
            );
          })}
          {steps.length > 4 && (
            <div className="w-5 h-5 rounded-full bg-gray-200 flex items-center justify-center text-[8px] text-gray-500 ring-2 ring-white">
              +{steps.length - 4}
            </div>
          )}
        </div>
        <span className="font-semibold ml-1">
          {steps.length} step{steps.length > 1 ? 's' : ''}
        </span>
        <svg
          className={cn('w-3 h-3 ml-auto transition-transform', expanded && 'rotate-180')}
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
        >
          <path d="M6 9l6 6 6-6" />
        </svg>
      </button>
      {expanded && (
        <div className="mt-1 space-y-1 animate-fade-in">
          {steps.map((step, i) => {
            const result = results[i];
            return (
              <div key={i} className="bg-white rounded-lg border border-gray-200/80 overflow-hidden shadow-soft">
                <div className={cn(
                  'px-2.5 py-1.5 border-b flex items-center gap-2',
                  result?.success ? 'bg-emerald-50/50 border-emerald-100/50' : 'bg-red-50/50 border-red-100/50'
                )}>
                  <div className={cn(
                    'w-4 h-4 rounded-full flex items-center justify-center text-[8px] text-white',
                    result?.success ? 'bg-emerald-500' : 'bg-red-500'
                  )}>
                    {result?.success ? '✓' : '✗'}
                  </div>
                  <span className="text-[11px] font-mono font-semibold text-gray-600">{step.action}</span>
                </div>
                <div className="px-2.5 py-1.5">
                  <p className="text-[11px] text-gray-500">{step.description}</p>
                  {result && (
                    <pre className="text-[9px] font-mono text-gray-400 mt-0.5 overflow-x-auto max-h-16 whitespace-pre-wrap break-words">
                      {result.output.length > 200 ? result.output.slice(0, 200) + '...' : result.output}
                    </pre>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

const ToolCallDisplay: React.FC<{
  toolCalls: ToolCall[];
  toolResults?: ToolResult[];
}> = ({ toolCalls, toolResults }) => {
  const [expanded, setExpanded] = useState(false);

  return (
    <div className="mb-2 w-full max-w-full min-w-0 animate-scale-in">
      <button
        onClick={() => setExpanded(!expanded)}
        className="flex items-center gap-2 px-2.5 py-1.5 rounded-lg bg-gray-50/80 border border-gray-200/80 text-[11px] text-gray-500 hover:bg-gray-100 transition-colors w-full"
      >
        <span className="text-[10px]">⚡</span>
        <span className="font-medium">
          {toolCalls.length} tool{toolCalls.length > 1 ? 's' : ''}
        </span>
        <svg
          className={cn('w-3 h-3 ml-auto transition-transform', expanded && 'rotate-180')}
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
        >
          <path d="M6 9l6 6 6-6" />
        </svg>
      </button>
      {expanded && (
        <div className="mt-1 space-y-1 animate-fade-in">
          {toolCalls.map((tc, i) => {
            const result = toolResults?.find((r) => r.tool_call_id === tc.id);
            return (
              <div key={tc.id} className="bg-white rounded-lg border border-gray-200/80 overflow-hidden shadow-soft">
                <div className="px-2.5 py-1 bg-gray-50/80 border-b border-gray-100">
                  <span className="text-[11px] font-mono font-semibold text-emerald-600">{tc.name}</span>
                </div>
                {result && (
                  <pre className="px-2.5 py-1.5 text-[10px] font-mono text-gray-500 overflow-x-auto max-h-24 whitespace-pre-wrap break-words">
                    {result.content.length > 300 ? result.content.slice(0, 300) + '...' : result.content}
                  </pre>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

export const TypingIndicator: React.FC<{ phase?: string }> = ({ phase = 'thinking' }) => {
  const phaseLabels: Record<string, string> = {
    thinking: 'Thinking...',
    planning: 'Planning steps...',
    executing: 'Working in Excel...',
    validating: 'Validating results...',
  };

  const phaseColors: Record<string, string> = {
    thinking: 'from-violet-400 to-violet-500',
    planning: 'from-blue-400 to-blue-500',
    executing: 'from-emerald-400 to-emerald-500',
    validating: 'from-amber-400 to-amber-500',
  };

  return (
    <div className="flex gap-2 mb-3 animate-slide-up">
      <div className={cn('flex-shrink-0 w-7 h-7 rounded-full bg-gradient-to-br flex items-center justify-center text-[10px] font-bold text-white shadow-sm', phaseColors[phase] || phaseColors.thinking)}>
        <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <path d="M12 2L2 7l10 5 10-5-10-5z" />
          <path d="M2 17l10 5 10-5" />
          <path d="M2 12l10 5 10-5" />
        </svg>
      </div>
      <div className="px-3.5 py-2.5 rounded-xl rounded-tl-sm bg-white border border-gray-200/80 shadow-soft">
        <p className="text-[11px] text-gray-400 font-medium mb-1">{phaseLabels[phase] || 'Processing...'}</p>
        <div className="flex gap-1">
          <div className="w-1.5 h-1.5 bg-gray-300 rounded-full animate-dot-pulse" style={{ animationDelay: '0ms' }} />
          <div className="w-1.5 h-1.5 bg-gray-300 rounded-full animate-dot-pulse" style={{ animationDelay: '150ms' }} />
          <div className="w-1.5 h-1.5 bg-gray-300 rounded-full animate-dot-pulse" style={{ animationDelay: '300ms' }} />
        </div>
      </div>
    </div>
  );
};

export const WelcomeScreen: React.FC = () => {
  return (
    <div className="flex-1 flex flex-col items-center justify-center px-8 py-12 animate-fade-in">
      <div className="relative mb-6">
        <div className="absolute inset-0 bg-gradient-to-br from-[#217346] to-[#185C37] rounded-3xl blur-lg opacity-30 scale-110" />
        <div className="w-20 h-20 rounded-3xl bg-gradient-to-br from-[#217346] to-[#185C37] flex items-center justify-center shadow-xl shadow-[#217346]/30 relative">
          <svg className="w-10 h-10 text-white" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M12 2L2 7l10 5 10-5-10-5z" />
            <path d="M2 17l10 5 10-5" />
            <path d="M2 12l10 5 10-5" />
          </svg>
        </div>
      </div>
      <h2 className="text-3xl font-bold text-gray-900 mb-2 tracking-tight">Cel</h2>
      <p className="text-base text-gray-600 text-center max-w-[320px] leading-relaxed font-medium">
        Tell me what to do.<br/>Consider it done.
      </p>
      <p className="text-xs text-gray-400 mt-4">Start typing below...</p>
    </div>
  );
};
