import React, { useState, useCallback, useRef } from 'react';
import { cn } from '@/lib/utils';

interface SelectedRangeInfo {
  address: string;
  sheetName: string;
  rowCount: number;
  columnCount: number;
  values: (string | number | boolean | null)[][];
}

interface ChatInputProps {
  onSend: (message: string) => void;
  disabled?: boolean;
  selectedRange?: SelectedRangeInfo | null;
}

export const ChatInput: React.FC<ChatInputProps> = ({ onSend, disabled, selectedRange }) => {
  const [input, setInput] = useState('');
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const handleSubmit = useCallback(() => {
    const trimmed = input.trim();
    if (!trimmed || disabled) return;
    onSend(trimmed);
    setInput('');
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
    }
  }, [input, disabled, onSend]);

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
    const el = e.target;
    el.style.height = 'auto';
    el.style.height = Math.min(el.scrollHeight, 100) + 'px';
  }, []);

  const handleCopy = useCallback(() => {
    const textarea = textareaRef.current;
    if (!textarea || !textarea.value) return;
    const selectedText = textarea.value.substring(textarea.selectionStart, textarea.selectionEnd);
    if (selectedText) {
      navigator.clipboard.writeText(selectedText);
    } else {
      navigator.clipboard.writeText(textarea.value);
    }
  }, []);

  const handleCut = useCallback(() => {
    const textarea = textareaRef.current;
    if (!textarea || !textarea.value || disabled) return;
    const selectedText = textarea.value.substring(textarea.selectionStart, textarea.selectionEnd);
    if (selectedText) {
      navigator.clipboard.writeText(selectedText);
      const newValue = textarea.value.substring(0, textarea.selectionStart) + textarea.value.substring(textarea.selectionEnd);
      setInput(newValue);
      if (textareaRef.current) {
        textareaRef.current.style.height = 'auto';
        textareaRef.current.style.height = Math.min(textareaRef.current.scrollHeight, 100) + 'px';
      }
    }
  }, [disabled]);

  const handleClear = useCallback(() => {
    if (disabled) return;
    setInput('');
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
    }
  }, [disabled]);

  const rangePreview = selectedRange
    ? `${selectedRange.address} (${selectedRange.sheetName} · ${selectedRange.rowCount}×${selectedRange.columnCount})`
    : null;

  return (
    <div className="flex-shrink-0 border-t border-gray-200/60 glass px-3 py-2.5">
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
            className={cn(
              'w-full bg-transparent text-[13px] text-gray-800 placeholder-gray-400 resize-none outline-none',
              'max-h-[100px] leading-relaxed font-normal'
            )}
            disabled={disabled}
          />
        </div>
        <div className="flex items-center gap-0.5 self-center">
          <button
            onClick={handleCopy}
            disabled={!input || disabled}
            title="Copy all"
            className="p-1.5 rounded-md text-gray-400 hover:text-[#217346] hover:bg-[#217346]/5 disabled:opacity-30 disabled:cursor-not-allowed transition-all"
          >
            <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <rect x="9" y="9" width="13" height="13" rx="2" />
              <path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1" />
            </svg>
          </button>
          <button
            onClick={handleCut}
            disabled={!input || disabled}
            title="Cut selected"
            className="p-1.5 rounded-md text-gray-400 hover:text-[#217346] hover:bg-[#217346]/5 disabled:opacity-30 disabled:cursor-not-allowed transition-all"
          >
            <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="6" cy="6" r="3" />
              <circle cx="6" cy="18" r="3" />
              <line x1="20" y1="4" x2="8.12" y2="15.88" />
              <line x1="14.47" y1="14.48" x2="20" y2="20" />
              <line x1="8.12" y1="8.12" x2="12" y2="12" />
            </svg>
          </button>
          <button
            onClick={handleClear}
            disabled={!input || disabled}
            title="Clear"
            className="p-1.5 rounded-md text-gray-400 hover:text-red-500 hover:bg-red-50 disabled:opacity-30 disabled:cursor-not-allowed transition-all"
          >
            <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <line x1="18" y1="6" x2="6" y2="18" />
              <line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </button>
        </div>
        <button
          onClick={handleSubmit}
          disabled={!input.trim() || disabled}
          className={cn(
            'flex-shrink-0 w-7 h-7 rounded-lg flex items-center justify-center transition-all duration-200 self-end',
            input.trim() && !disabled
              ? 'bg-gradient-to-br from-[#217346] to-[#185C37] text-white shadow-green hover:shadow-md active:scale-95'
              : 'bg-gray-200/80 text-gray-400 cursor-not-allowed'
          )}
        >
          <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
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
