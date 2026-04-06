import React, { useState, useCallback, useRef, useEffect } from 'react';
import { cn } from '@/lib/utils';

interface SelectedRangeInfo {
  address: string;
  sheetName: string;
  rowCount: number;
  columnCount: number;
  values: (string | number | boolean | null)[][];
}

interface ChatInputProps {
  onSend: (message: string, enableReasoning?: boolean) => void;
  disabled?: boolean;
  selectedRange?: SelectedRangeInfo | null;
  reasoningEnabled?: boolean;
  onToggleReasoning?: (enabled: boolean) => void;
}

export const ChatInput: React.FC<ChatInputProps> = ({ 
  onSend, 
  disabled, 
  selectedRange,
  reasoningEnabled = true,
  onToggleReasoning 
}) => {
  const [input, setInput] = useState('');
  const [reasoning, setReasoning] = useState(reasoningEnabled);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    setReasoning(reasoningEnabled);
  }, [reasoningEnabled]);

  const handleSubmit = useCallback(() => {
    const trimmed = input.trim();
    if (!trimmed || disabled) return;
    onSend(trimmed, reasoning);
    setInput('');
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
    }
  }, [input, disabled, onSend, reasoning]);

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        handleSubmit();
      }
    },
    [handleSubmit]
  );

  const handleInput = useCallback((e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setInput(e.target.value);
    requestAnimationFrame(() => {
      const el = e.target;
      el.style.height = 'auto';
      el.style.height = Math.min(el.scrollHeight, 120) + 'px';
    });
  }, []);

  const handleCancel = useCallback(() => {
    setInput('');
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
    }
  }, []);

  const toggleReasoning = useCallback(() => {
    const newValue = !reasoning;
    setReasoning(newValue);
    onToggleReasoning?.(newValue);
  }, [reasoning, onToggleReasoning]);

  const rangePreview = selectedRange
    ? `${selectedRange.address} (${selectedRange.sheetName} · ${selectedRange.rowCount}×${selectedRange.columnCount})`
    : null;

  return (
    <div className="flex-shrink-0 border-t border-gray-200/60 glass px-3 py-2.5">
      {/* Reasoning toggle row */}
      <div className="flex items-center justify-between mb-2">
        <button
          onClick={toggleReasoning}
          className={cn(
            'flex items-center gap-1.5 px-2 py-1 rounded-md text-[10px] font-medium transition-all',
            reasoning 
              ? 'bg-[#217346]/10 text-[#217346] border border-[#217346]/20' 
              : 'bg-gray-100 text-gray-400 border border-gray-200'
          )}
        >
          <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M12 2a10 10 0 1 0 10 10H12V2z" />
            <path d="M12 2a10 10 0 0 1 10 10" />
            <circle cx="12" cy="12" r="3" />
          </svg>
          {reasoning ? 'Think' : 'Fast'}
        </button>
        <span className="text-[9px] text-gray-400">
          {reasoning ? 'Detailed thinking on' : 'Quick response'}
        </span>
      </div>
      
      <div className="flex items-end gap-2 bg-gray-50/80 rounded-xl border border-gray-200/80 focus-within:border-[#217346]/40 focus-within:ring-2 focus-within:ring-[#217346]/10 transition-all duration-200 px-3 py-2 shadow-soft">
        <div className="flex-1 min-w-0">
          {rangePreview && (
            <div className="flex items-center gap-1.5 mb-1.5 px-2 py-1 rounded-md bg-[#217346]/5 border border-[#217346]/10 animate-fade-in">
              <svg className="w-3 h-3 text-[#217346] flex-shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <rect x="3" y="3" width="18" height="18" rx="2" />
                <path d="M3 9h18M3 15h18M9 3v18M15 3v18" />
              </svg>
              <span className="text-[10px] font-semibold text-[#217346] truncate">{rangePreview}</span>
            </div>
          )}
          <textarea
            ref={textareaRef}
            value={input}
            onChange={handleInput}
            onKeyDown={handleKeyDown}
            placeholder={rangePreview ? "Ask about selected range..." : "Ask Cel to do something..."}
            rows={1}
            style={{ height: 'auto' }}
            className={cn(
              'w-full bg-transparent text-[13px] text-gray-800 placeholder-gray-400 resize-none outline-none',
              'max-h-[120px] leading-relaxed font-normal'
            )}
            disabled={disabled}
          />
        </div>
        
        {input && !disabled && (
          <button
            onClick={handleCancel}
            title="Clear"
            className="p-1.5 rounded-md text-gray-400 hover:text-red-500 hover:bg-red-50 transition-all"
          >
            <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <line x1="18" y1="6" x2="6" y2="18" />
              <line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </button>
        )}
        <button
          onClick={handleSubmit}
          disabled={!input.trim() || disabled}
          className={cn(
            'flex-shrink-0 w-8 h-8 rounded-lg flex items-center justify-center transition-all duration-200 self-end',
            input.trim() && !disabled
              ? 'bg-gradient-to-br from-[#217346] to-[#185C37] text-white shadow-green hover:shadow-md active:scale-95'
              : 'bg-gray-200/80 text-gray-400 cursor-not-allowed'
          )}
        >
          <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
            <path d="M5 12h14M12 5l7 7-7 7" />
          </svg>
        </button>
      </div>
      <div className="flex items-center justify-center gap-1 mt-1.5">
        <div className="w-1 h-1 rounded-full bg-[#217346]/30" />
        <p className="text-[9px] text-gray-400 font-medium tracking-wide uppercase">
          Cel · AI can make mistakes
        </p>
        <div className="w-1 h-1 rounded-full bg-[#217346]/30" />
      </div>
    </div>
  );
};
