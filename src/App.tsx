import React, { useState, useCallback, useRef, useEffect } from 'react';
import { cn } from '@/lib/utils';
import { MessageBubble, TypingIndicator, WelcomeScreen } from '@/components/MessageBubble';
import { ChatInput } from '@/components/ChatInput';
import { SettingsPanel } from '@/components/SettingsPanel';
import { useConfirm, ConfirmModal } from '@/components/ConfirmModal';
import type { ChatMessage, AIProvider } from '@/lib/ai-providers';
import * as ExcelAPI from '@/lib/excel-api';
import { loadMemory, saveMemory, addMistake, trackOperation, addCheckpoint, getCheckpoints, removeCheckpoint, getMemoryContext } from '@/lib/memory';

type ProcessingPhase =
  | 'idle'
  | 'understanding'
  | 'context'
  | 'schema'
  | 'intent'
  | 'thinking'
  | 'reasoning'
  | 'planning'
  | 'analyzing'
  | 'confirming'
  | 'executing'
  | 'verifying'
  | 'validating';

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

interface ExtendedMessage extends ChatMessage {
  id: string;
  planSteps?: { action: string; params: Record<string, any>; description: string }[];
  executionResults?: { action: string; success: boolean; output: string }[];
  chartImage?: string;
  chartImages?: string[];
  isReasoning?: boolean;
}

interface SelectedRangeInfo {
  address: string;
  sheetName: string;
  rowCount: number;
  columnCount: number;
  values: (string | number | boolean | null)[][];
}

const PHASE_CONFIG: Record<ProcessingPhase, { label: string; color: string; accent: string }> = {
  idle: { label: '', color: '', accent: '' },
  understanding: { label: 'Understanding your request...', color: 'text-violet-600', accent: 'from-violet-500 to-purple-600' },
  context: { label: 'Reading selected range and context...', color: 'text-indigo-600', accent: 'from-indigo-500 to-blue-600' },
  schema: { label: 'Analyzing table schema...', color: 'text-cyan-600', accent: 'from-cyan-500 to-blue-600' },
  intent: { label: 'Mapping intent to columns...', color: 'text-blue-600', accent: 'from-blue-500 to-indigo-600' },
  thinking: { label: 'Analyzing...', color: 'text-violet-600', accent: 'from-violet-500 to-purple-600' },
  reasoning: { label: 'Thinking...', color: 'text-indigo-600', accent: 'from-indigo-500 to-purple-600' },
  planning: { label: 'Planning...', color: 'text-blue-600', accent: 'from-blue-500 to-indigo-600' },
  analyzing: { label: 'Analyzing data...', color: 'text-cyan-600', accent: 'from-cyan-500 to-blue-600' },
  confirming: { label: 'Confirm...', color: 'text-orange-600', accent: 'from-orange-500 to-red-600' },
  executing: { label: 'Working...', color: 'text-emerald-600', accent: 'from-emerald-500 to-teal-600' },
  verifying: { label: 'Verifying chart mappings...', color: 'text-amber-600', accent: 'from-amber-500 to-orange-600' },
  validating: { label: 'Validating...', color: 'text-amber-600', accent: 'from-amber-500 to-orange-600' },
};

export default function App() {
  const [messages, setMessages] = useState<ExtendedMessage[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingPhase, setProcessingPhase] = useState<ProcessingPhase>('idle');
  const [showSettings, setShowSettings] = useState(false);
  const [showScrollDown, setShowScrollDown] = useState(false);
  const [reasoningEnabled, setReasoningEnabled] = useState(true);
  const [selectedRange, setSelectedRange] = useState<SelectedRangeInfo | null>(null);
  const [executionProgress, setExecutionProgress] = useState<{ current: number; total: number; action: string } | null>(null);
  const [settings, setSettings] = useState<Settings>(() => {
    try {
      const saved = localStorage.getItem('excel-ai-settings');
      if (saved) return JSON.parse(saved);
    } catch {}
    return { provider: 'openai' as const, model: '', openaiKey: '', anthropicKey: '', googleKey: '' };
  });

  // Custom confirm modal for Office.js compatibility
  const { confirm, Modal: ConfirmModalComponent } = useConfirm();

  const messagesEndRef = useRef<HTMLDivElement>(null);
  const messagesRef = useRef<ExtendedMessage[]>([]);
  messagesRef.current = messages;

  // Poll for selected range changes
  useEffect(() => {
    const poll = async () => {
      try {
        if (typeof Excel === 'undefined') return;
        const data = await getSelectedRangeData();
        if (data) {
          const parsed = JSON.parse(data);
          if (parsed.address && parsed.sheetName) {
            setSelectedRange(parsed);
            return;
          }
        }
        setSelectedRange(null);
      } catch {
        setSelectedRange(null);
      }
    };

    poll();
    const interval = setInterval(poll, 2000);
    return () => clearInterval(interval);
  }, []);

  const scrollToBottom = useCallback(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, []);

  const handleUndo = useCallback(async () => {
    const checkpoints = getCheckpoints();
    if (checkpoints.length === 0) return;
    const latest = checkpoints[0]; // Most recent is first
    try {
      await ExcelAPI.setValues(latest.address, latest.values, latest.sheet);
      removeCheckpoint(latest.id);
      setMessages((prev) => [...prev, {
        id: `ai-${Date.now()}`,
        role: 'assistant',
        content: `Restored ${latest.sheet}!${latest.address} to its previous state (before ${latest.description}).`,
      }]);
    } catch (err: any) {
      setMessages((prev) => [...prev, {
        id: `ai-${Date.now()}`,
        role: 'assistant',
        content: `Failed to restore checkpoint: ${err.message}`,
      }]);
    }
  }, []);

  const getApiKey = useCallback((): string => {
    switch (settings.provider) {
      case 'openai': return settings.openaiKey;
      case 'anthropic': return settings.anthropicKey;
      case 'google': return settings.googleKey;
      case 'openrouter': return settings.openrouterKey;
      default: return '';
    }
  }, [settings]);

  const processMessage = useCallback(
    async (userMessage: string) => {
      const apiKey = getApiKey();
      if (!apiKey) {
        setMessages((prev) => [
          ...prev,
          { id: `user-${Date.now()}`, role: 'user', content: userMessage },
          { id: `msg-${Date.now()}`, role: 'assistant', content: 'Please configure your API key in Settings to get started.' },
        ]);
        return;
      }

      const userMsg: ExtendedMessage = { id: `user-${Date.now()}`, role: 'user', content: userMessage };
      setMessages((prev) => [...prev, userMsg]);
      setIsProcessing(true);
      setProcessingPhase('understanding');

      try {
        setProcessingPhase('context');
        const workbookContext = await getWorkbookContext();
        const selectedRangeData = await getSelectedRangeData();

        // Resource-aware: Use user's model as-is - truly dynamic, works with ANY model
        const effectiveModel = settings.model;

        setProcessingPhase('thinking');
        
        // Create placeholder messages for streaming
        let plan: any[] = [];
        let reasoning = '';
        let responseContent = '';
        
        const response = await fetch('/api/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            message: userMessage,
            provider: settings.provider,
            model: effectiveModel,
            api_key: apiKey,
            workbook_context: workbookContext,
            selected_range: selectedRangeData,
            memory_context: getMemoryContext(),
            enable_reasoning: reasoningEnabled,
            conversation_history: messagesRef.current
              .filter(m => m.role === 'user' || m.role === 'assistant')
              .slice(-6)  // Last 3 exchanges
              .map(m => ({ role: m.role, content: m.content.slice(0, 500) })),
          }),
        });

        if (!response.ok) {
          const text = await response.text();
          throw new Error(`Backend error: ${response.status} - ${text}`);
        }

        // Handle streaming response
        const reader = response.body?.getReader();
        const decoder = new TextDecoder();
        
        if (!reader) {
          throw new Error('No response body');
        }

        let buffer = '';
        let reasoningMessageId: string | null = null;
        
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          
          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          buffer = lines.pop() || '';
          
          for (const line of lines) {
            if (line.startsWith('data: ')) {
              try {
                const data = JSON.parse(line.slice(6));
                
                if (data.type === 'reasoning' && data.content) {
                  setProcessingPhase('reasoning');
                  reasoning = data.content;
                  
                  const cleanReasoning = reasoning
                    .replace(/\*/g, '')
                    .replace(/^\d+\.\s*/gm, '')
                    .replace(/^-\s*/gm, '')
                    .replace(/^•\s*/gm, '')
                    .replace(/\n{2,}/g, '\n')
                    .trim();
                  
                  if (!reasoningMessageId) {
                    reasoningMessageId = `reasoning-${Date.now()}`;
                    const reasoningMsg: ExtendedMessage = {
                      id: reasoningMessageId,
                      role: 'assistant',
                      content: cleanReasoning,
                      isReasoning: true,
                    };
                    setMessages((prev) => [...prev, reasoningMsg]);
                  } else {
                    setMessages((prev) => prev.map(m => 
                      m.id === reasoningMessageId ? { ...m, content: cleanReasoning } : m
                    ));
                  }
                  setTimeout(scrollToBottom, 100);
                }
                
                if (data.type === 'plan' && data.content !== undefined) {
                  plan = data.plan || [];
                  responseContent = data.content;
                  setTimeout(scrollToBottom, 100);
                }

                if (data.type === 'clarification' && data.content) {
                  // Clarification agent detected ambiguity — show question, stop processing
                  setMessages((prev) => [...prev, {
                    id: `clarify-${Date.now()}`,
                    role: 'assistant',
                    content: data.content,
                  }]);
                  setIsProcessing(false);
                  setProcessingPhase('idle');
                  return;
                }
              } catch {}
            }
          }
        }

        // Get final plan from the last message
        const lastMsg = messages[messages.length - 1];
        plan = plan.length > 0 ? plan : [];
        
        const { validSteps, validationErrors } = validatePlan(plan, userMessage);

        // Parse selected range data for context awareness
        let selectedRange: { address: string; sheetName: string; rowCount: number; columnCount: number } | null = null;
        if (selectedRangeData) {
          try {
            const sr = JSON.parse(selectedRangeData);
            console.log('[DEBUG] selectedRangeData parsed:', sr);
            if (sr.address && sr.sheetName) {
              selectedRange = { address: sr.address, sheetName: sr.sheetName, rowCount: sr.rowCount || 1, columnCount: sr.columnCount || 1 };
              console.log('[DEBUG] selectedRange set:', selectedRange);
            }
          } catch (e) { console.log('[DEBUG] selectedRangeData parse error:', e); }
        }

        // If user has a selection AND is modifying data (not creating new sheet), use the selection
        const hasSelection = selectedRange && selectedRange.address;
        const createsNewSheet = /new sheet|another sheet|make sheet|create sheet/i.test(userMessage);
        const modifiesExisting = /add|write|put|fill|update|change|modify|calculate|compute|less|more|subtract|increase|decrease/i.test(userMessage);
        const userWantsSelected = hasSelection && !createsNewSheet && (modifiesExisting || /selected|here|this range|these cells|border/i.test(userMessage));
        console.log('[DEBUG] userWantsSelected:', userWantsSelected, '| hasSelection:', hasSelection, '| createsNewSheet:', createsNewSheet);
        console.log('[DEBUG] userWantsSelected regex test for:', userMessage, '→ result:', userWantsSelected);

        // Check if user wants data generation
        const dataGenMatch = userMessage.match(/create\s+(a\s+)?table.*?(?:with|for)\s+(\d+)\s+(people|records|employees|sales|students|rows)/i);
        const newSheetMatch = /new\s+sheet|make\s+sheet|create\s+sheet|another\s+sheet/i.test(userMessage);
        const bloodReportMatch = /blood\s*report|blood\s*test|medical\s*report/i.test(userMessage);
        const employeeMatch = /employee|staff|worker|team\s*member/i.test(userMessage);
        const salesMatch = /sales|revenue|product|order/i.test(userMessage);
        const studentMatch = /student|grade|class|exam|mark/i.test(userMessage);

        if (dataGenMatch) {
          const count = parseInt(dataGenMatch[2]);
          let dataType = 'blood_report';
          if (employeeMatch) dataType = 'employee';
          else if (salesMatch) dataType = 'sales';
          else if (studentMatch) dataType = 'student';
          else if (bloodReportMatch) dataType = 'blood_report';

          setProcessingPhase('analyzing');
          const genResponse = await fetch('/api/generate-data', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ data_type: dataType, count: Math.min(count, 10000) }),
          });

          if (genResponse.ok) {
            const genResult = await genResponse.json();
            const values = genResult.data;
            const rows = values.length;
            const cols = values[0]?.length || 1;
            const colLetter = cols <= 26 ? String.fromCharCode(64 + cols) : String.fromCharCode(64 + Math.floor((cols - 1) / 26)) + String.fromCharCode(64 + ((cols - 1) % 26) + 1);
            const address = `A1:${colLetter}${rows}`;

            try {
              let targetSheet = 'Sheet1';
              if (newSheetMatch) {
                const sheetName = `${dataType.charAt(0).toUpperCase() + dataType.slice(1)}_${Date.now().toString().slice(-6)}`;
                await ExcelAPI.addWorksheet(sheetName);
                targetSheet = sheetName;
              }
              
              await ExcelAPI.setValues(address, values, targetSheet);
              await ExcelAPI.createTable(address, `Table_${Date.now()}`, targetSheet);
              
              const aiMsg: ExtendedMessage = {
                id: `msg-${Date.now()}`,
                role: 'assistant',
                content: `✅ Created table with ${count} ${dataType.replace('_', ' ')} records on sheet "${targetSheet}"!\n\n📊 Data: ${address} (${rows} rows × ${cols} columns)\n📋 Columns: ${genResult.headers.join(', ')}`,
              };
              setMessages((prev) => [...prev, aiMsg]);
              setIsProcessing(false);
              setProcessingPhase('idle');
              return;
            } catch (err: any) {
              const aiMsg: ExtendedMessage = {
                id: `msg-${Date.now()}`,
                role: 'assistant',
                content: `Generated ${count} records but couldn't write to Excel: ${err.message}`,
              };
              setMessages((prev) => [...prev, aiMsg]);
              setIsProcessing(false);
              setProcessingPhase('idle');
              return;
            }
          }
        }

        // Check if user wants data analysis
        const analysisKeywords = /analyze|analysis|statistics|stats|trends|insights|outliers|distribution|compare/i;
        const chartKeywords = /chart|graph|plot|visual|pie|bar|line|trend/i;
        const isAnalysisRequest = analysisKeywords.test(userMessage);
        const isChartRequest = chartKeywords.test(userMessage);
        const isPieRequest = /pie\s*chart/i.test(userMessage);
        const isBarRequest = /bar\s*(graph|chart)/i.test(userMessage);
        const isLineRequest = /line\s*chart|trend/i.test(userMessage);

        if ((isAnalysisRequest || isChartRequest) && selectedRangeData) {
          setProcessingPhase('context');
          let selectedData = null;
          try {
            const sr = JSON.parse(selectedRangeData);
            if (sr.values && sr.values.length > 0) {
              selectedData = sr.values;
            }
          } catch {}

          if (selectedData) {
            setProcessingPhase('schema');
            try {
              await fetch('/api/build-schema', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                  data: selectedData,
                  question: userMessage,
                  provider: settings.provider,
                  model: effectiveModel,
                  api_key: apiKey,
                  headers: selectedData[0] || [],
                }),
              });
            } catch {}

            let analysisResult: any = { analysis: '' };
            if (isAnalysisRequest) {
              setProcessingPhase('analyzing');
              const isLargeData = selectedData.length > 1000;
              const endpoint = isLargeData ? '/api/analyze-large' : '/api/analyze';
              const analyzeResponse = await fetch(endpoint, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                  data: selectedData,
                  question: userMessage,
                  provider: settings.provider,
                  model: effectiveModel,
                  api_key: apiKey,
                  headers: selectedData[0] || [],
                }),
              });
              if (analyzeResponse.ok) {
                analysisResult = await analyzeResponse.json();
              }
            }

            if (isChartRequest && selectedRangeData) {
                try {
                  const sr = JSON.parse(selectedRangeData);
                  let headers = selectedData[0] || [];
                  let rows = selectedData.slice(1);
                  let address = sr.address;
                  const sheetName = sr.sheetName;
                  let chartsCreated = 0;
                  let chartDesc = '';

                  const userMsg = userMessage.toLowerCase();

                  const addrParts = address.split('!');
                  const rangePart = addrParts.length > 1 ? addrParts[1] : addrParts[0];
                  const rangeCells = rangePart.split(':');
                  const startCell = rangeCells[0];
                  const endCell = rangeCells.length > 1 ? rangeCells[1] : rangeCells[0];
                  const startColMatch = startCell.match(/^([A-Z]+)/);
                  const startColLabel = startColMatch ? startColMatch[1] : 'A';
                  const colLabelToIndex = (label: string): number => {
                    return label.toUpperCase().split('').reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
                  };
                  const indexToColLabel = (idx: number): string => {
                    let n = idx + 1;
                    let out = '';
                    while (n > 0) {
                      const rem = (n - 1) % 26;
                      out = String.fromCharCode(65 + rem) + out;
                      n = Math.floor((n - 1) / 26);
                    }
                    return out;
                  };
                  const startColIdx = colLabelToIndex(startColLabel);
                  const numCols = headers.length;

                  // Build column profiles from the selected range (dynamic for any table shape).
                  const numericCols: number[] = [];
                  const categoricalCols: number[] = [];
                  const colScores: number[] = Array(headers.length).fill(0);
                  const headerNorms = headers.map((h: any) => String(h ?? '').trim().toLowerCase());
                  const tokenize = (s: string) => s.toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
                  
                  for (let i = 0; i < headers.length; i++) {
                    let numCount = 0;
                    let catCount = 0;
                    const sampleSize = Math.min(10, rows.length);
                    
                    for (let j = 0; j < sampleSize; j++) {
                      const val = rows[j]?.[i];
                      if (val === undefined || val === null || val === '') continue;
                      const strVal = String(val).replace(/[^0-9.-]/g, '');
                      if (!isNaN(Number(strVal)) && strVal !== '') {
                        numCount++;
                      } else {
                        catCount++;
                      }
                    }
                    
                    if (numCount > catCount) numericCols.push(i);
                    else categoricalCols.push(i);
                  }

                  // Dynamic intent extraction: explicit matches > phrase-level mapping > fuzzy token fallback.
                  let userMentionedCols: number[] = [];
                  const userMsgLower = userMessage.toLowerCase();
                  const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                  
                  // Extract column names from user message - remove common words
                  const skipWords = ['bar', 'graph', 'chart', 'pie', 'line', 'make', 'create', 'show', 'display', 'the', 'a', 'an', 'of', 'to', 'for', 'between', 'with', 'on', 'in', 'and', 'vs', 'versus', 'column', 'columns'];
                  const userWords = userMsgLower.split(/\s+/).filter(w => !skipWords.includes(w) && w.length > 1);
                  const mentionedSet = new Set<number>();

                  // Strongest signal: exact header names in the prompt.
                  // If user says "product and quantity", we should lock onto those columns.
                  const explicitHeaderSet = new Set<number>();
                  for (let i = 0; i < headers.length; i++) {
                    const rawHeader = String(headers[i] ?? '').trim().toLowerCase();
                    if (!rawHeader) continue;
                    const exactHeaderRegex = new RegExp(`\\b${escapeRegExp(rawHeader)}\\b`, 'i');
                    if (exactHeaderRegex.test(userMsgLower)) {
                      explicitHeaderSet.add(i);
                      mentionedSet.add(i);
                      colScores[i] += 5;
                    }
                  }

                  // Handle explicit Excel column letters only in explicit column-reference phrases.
                  // Examples: "column b and d", "b, d columns", "columns aa and ac".
                  const explicitColumnTokens = new Set<string>();
                  const forwardPattern = /\bcolumn[s]?\s+([a-z]{1,3}(?:\s*(?:,|and|&)\s*[a-z]{1,3})*)\b/gi;
                  const reversePattern = /\b([a-z]{1,3}(?:\s*(?:,|and|&)\s*[a-z]{1,3})+)\s+column[s]?\b/gi;

                  const collectColumnTokens = (groupText: string) => {
                    const tokens = groupText
                      .split(/\s*(?:,|and|&)\s*/i)
                      .map((t) => t.trim().toLowerCase())
                      .filter((t) => /^[a-z]{1,3}$/.test(t));
                    for (const t of tokens) explicitColumnTokens.add(t);
                  };

                  let patternMatch: RegExpExecArray | null;
                  while ((patternMatch = forwardPattern.exec(userMsgLower)) !== null) {
                    collectColumnTokens(patternMatch[1]);
                  }
                  while ((patternMatch = reversePattern.exec(userMsgLower)) !== null) {
                    collectColumnTokens(patternMatch[1]);
                  }

                  for (const token of explicitColumnTokens) {
                    const absoluteIdx = colLabelToIndex(token.toUpperCase());
                    const relativeIdx = absoluteIdx - startColIdx;
                    if (relativeIdx >= 0 && relativeIdx < headers.length) {
                      mentionedSet.add(relativeIdx);
                      colScores[relativeIdx] += 4;
                    }
                  }

                  const phraseMatch = userMsgLower.match(
                    /(?:for|plot|graph|chart)\s+(.+?)(?:\s+(?:and|vs|versus|by)\s+)(.+?)(?:$|[,.!?])/
                  );
                  const requestedTerms: string[] = [];
                  if (phraseMatch) {
                    requestedTerms.push(phraseMatch[1].trim(), phraseMatch[2].trim());
                  }

                  const scoreHeaderMatch = (header: string, term: string): number => {
                    if (!header || !term) return 0;
                    const h = header.toLowerCase().trim();
                    const t = term.toLowerCase().trim();
                    if (h === t) return 1.0;
                    if (h.includes(t) || t.includes(h)) return 0.85;
                    const hTokens = tokenize(h);
                    const tTokens = tokenize(t);
                    if (!hTokens.length || !tTokens.length) return 0;
                    const overlap = hTokens.filter((tok) => tTokens.includes(tok)).length;
                    if (!overlap) return 0;
                    return overlap / Math.max(hTokens.length, tTokens.length);
                  };

                  const resolveBestColForTerm = (term: string): number | null => {
                    let bestIdx: number | null = null;
                    let bestScore = 0;
                    for (let i = 0; i < headerNorms.length; i++) {
                      const s = scoreHeaderMatch(headerNorms[i], term);
                      if (s > bestScore) {
                        bestScore = s;
                        bestIdx = i;
                      }
                    }
                    return bestScore >= 0.55 ? bestIdx : null;
                  };

                  for (const term of requestedTerms) {
                    const idx = resolveBestColForTerm(term);
                    if (idx !== null) {
                      mentionedSet.add(idx);
                      colScores[idx] += 3;
                    }
                  }
                  
                  // Match each user word against actual headers
                  for (let i = 0; i < headers.length; i++) {
                    const h = headers[i].toLowerCase();
                    for (const userWord of userWords) {
                      // Fuzzy match: header contains word OR word matches part of header
                      if (h.includes(userWord) || userWord.includes(h) || 
                          h.split(/\s+/).some((hw: string) => hw.startsWith(userWord.slice(0, 3)) || userWord.startsWith(hw.slice(0, 3)))) {
                        mentionedSet.add(i);
                        colScores[i] += 1;
                        break;
                      }
                    }
                  }
                  userMentionedCols = Array.from(mentionedSet).sort((a, b) => colScores[b] - colScores[a]);

                  // Dynamic AI-recommended column extraction - match column names from analysis
                  let aiRecommendedCols: number[] = [];
                  if (analysisResult.analysis) {
                    const analysisText = analysisResult.analysis.toLowerCase();
                    
                    // Look for "X vs Y" or "X and Y" patterns for chart recommendations
                    const vsMatch = analysisText.match(/(?:vs|versus|and|by)\s+([a-z\s]+?)(?:\s+(?:would|make|good|best)|$)/i);
                    if (vsMatch) {
                      const targetCol = vsMatch[1].trim().toLowerCase();
                      for (let i = 0; i < headers.length; i++) {
                        const h = headers[i].toLowerCase();
                        if (h.includes(targetCol) || targetCol.includes(h)) {
                          aiRecommendedCols.push(i);
                        }
                      }
                    }
                    
                    // If no vs pattern, find any column names mentioned in analysis
                    if (aiRecommendedCols.length < 2) {
                      for (let i = 0; i < headers.length; i++) {
                        const h = headers[i].toLowerCase();
                        const words = h.split(/\s+/).filter((w: string) => w.length > 2);
                        for (const word of words) {
                          if (analysisText.includes(word) && !aiRecommendedCols.includes(i)) {
                            aiRecommendedCols.push(i);
                            break;
                          }
                        }
                        if (aiRecommendedCols.length >= 2) break;
                      }
                    }
                  }

                  // Backend-authoritative intent resolution (scalable source of truth).
                  setProcessingPhase('intent');
                  let backendResolvedCols: number[] = [];
                  let backendChartRequests: { chart_type: string; x_col_index: number; y_col_index: number; x_header?: string; y_header?: string }[] = [];
                  let backendNeedsClarification = false;
                  let backendClarificationReason = '';
                  let backendClarificationQuestion = '';
                  try {
                    const intentRes = await fetch('/api/resolve-chart-intent', {
                      method: 'POST',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({
                        data: selectedData,
                        question: userMessage,
                        headers,
                        start_col_index: startColIdx,
                      }),
                    });
                    if (intentRes.ok) {
                      const intent = await intentRes.json();
                      if (Array.isArray(intent.chart_requests)) {
                        backendChartRequests = intent.chart_requests
                          .filter((r: any) => typeof r?.x_col_index === 'number' && typeof r?.y_col_index === 'number')
                          .map((r: any) => ({
                            chart_type: typeof r?.chart_type === 'string' ? r.chart_type : 'column',
                            x_col_index: r.x_col_index,
                            y_col_index: r.y_col_index,
                            x_header: r.x_header,
                            y_header: r.y_header,
                          }));
                      }
                      if (typeof intent.x_col_index === 'number' && typeof intent.y_col_index === 'number') {
                        backendResolvedCols = [intent.x_col_index, intent.y_col_index];
                      }
                      backendNeedsClarification =
                        !!intent.needs_clarification &&
                        backendResolvedCols.length < 2 &&
                        backendChartRequests.length === 0;
                      backendClarificationReason = intent.reason || '';
                      backendClarificationQuestion = intent.clarification_question || '';
                    }
                  } catch {}

                  if (backendNeedsClarification) {
                    const aiMsg: ExtendedMessage = {
                      id: `msg-${Date.now()}`,
                      role: 'assistant',
                      content: `${analysisResult.analysis || 'Analysis complete.'}\n\nI need a quick clarification before charting: ${backendClarificationQuestion || backendClarificationReason || 'Please confirm which two columns to use (for example: "Product and Quantity" or "columns B and D").'}`,
                    };
                    setMessages((prev) => [...prev, aiMsg]);
                    setIsProcessing(false);
                    setProcessingPhase('idle');
                    return;
                  }

                  setProcessingPhase('planning');
                  const startRow = startCell.match(/\d+/)?.[0] || '1';
                  const endRow = endCell.match(/(\d+)$/)?.[1] || '100';

                  const createChartForColumn = async (colIdx: number, chartType: string, title: string) => {
                    const colLetter = indexToColLabel(startColIdx + colIdx);
                    const chartRange = `${colLetter}${startRow}:${colLetter}${endRow}`;
                    await ExcelAPI.createChart(chartType, chartRange, sheetName, title);
                    chartsCreated++;
                    chartDesc += `📊 ${chartType} chart: ${title}\n`;
                  };

                  const createChartWithTwoColumns = async (col1Idx: number, col2Idx: number, chartType: string, title: string) => {
                    const firstLetter = indexToColLabel(startColIdx + col1Idx);
                    const secondLetter = indexToColLabel(startColIdx + col2Idx);
                    const firstRange = `${firstLetter}${startRow}:${firstLetter}${endRow}`;
                    const secondRange = `${secondLetter}${startRow}:${secondLetter}${endRow}`;
                    await ExcelAPI.createChartFromTwoColumns(chartType, firstRange, secondRange, sheetName, title);
                    chartsCreated++;
                    chartDesc += `📊 ${chartType} chart: ${title}\n`;
                  };

                  const createdChartMappings: { x_header: string; y_header: string; chart_type: string }[] = [];
                  setProcessingPhase('executing');

                  if (backendChartRequests.length > 0) {
                    for (const req of backendChartRequests) {
                      const xIdx = req.x_col_index;
                      const yIdx = req.y_col_index;
                      if (xIdx < 0 || yIdx < 0 || xIdx >= headers.length || yIdx >= headers.length || xIdx === yIdx) continue;
                      const chartType = req.chart_type || (isPieRequest ? 'pie' : isLineRequest ? 'line' : 'column');
                      const xHeader = req.x_header || headers[xIdx];
                      const yHeader = req.y_header || headers[yIdx];
                      const title =
                        chartType === 'line'
                          ? `${yHeader} Trend by ${xHeader}`
                          : `${yHeader} by ${xHeader}`;
                      await createChartWithTwoColumns(xIdx, yIdx, chartType, title);
                      createdChartMappings.push({ x_header: String(xHeader), y_header: String(yHeader), chart_type: chartType });
                    }
                  }

                  // Priority contract:
                  // 1) explicit header mentions, 2) high-confidence user intent mapping, 3) AI fallback.
                  const explicitHeaderCols = Array.from(explicitHeaderSet);
                  const confidentUserCols = userMentionedCols.filter((c) => colScores[c] >= 3);
                  const effectiveCols =
                    backendResolvedCols.length >= 2 ? backendResolvedCols :
                    explicitHeaderCols.length >= 2 ? explicitHeaderCols :
                    confidentUserCols.length >= 2 ? confidentUserCols :
                    userMentionedCols.length >= 2 ? userMentionedCols :
                    aiRecommendedCols.length >= 2 ? aiRecommendedCols : [];

                  // If user clearly asked for two fields but we couldn't resolve both confidently,
                  // ask for clarification instead of guessing wrong columns.
                  if (requestedTerms.length >= 2 && effectiveCols.length < 2) {
                    const aiMsg: ExtendedMessage = {
                      id: `msg-${Date.now()}`,
                      role: 'assistant',
                      content: `I could not confidently map both columns from "${requestedTerms[0]}" and "${requestedTerms[1]}". Please confirm the two columns (for example: "Product and Quantity" or "columns B and D").`,
                    };
                    setMessages((prev) => [...prev, aiMsg]);
                    setIsProcessing(false);
                    setProcessingPhase('idle');
                    return;
                  }

                  if (chartsCreated === 0 && effectiveCols.length >= 2) {
                    const catCol = effectiveCols.find(c => categoricalCols.includes(c)) ?? effectiveCols[0];
                    const numCol = effectiveCols.find(c => numericCols.includes(c)) ?? effectiveCols[1];
                    if (isPieRequest) {
                      await createChartWithTwoColumns(catCol, numCol, 'pie', `${headers[catCol]} vs ${headers[numCol]}`);
                      createdChartMappings.push({ x_header: String(headers[catCol]), y_header: String(headers[numCol]), chart_type: 'pie' });
                    } else if (isBarRequest) {
                      await createChartWithTwoColumns(catCol, numCol, 'column', `${headers[numCol]} by ${headers[catCol]}`);
                      createdChartMappings.push({ x_header: String(headers[catCol]), y_header: String(headers[numCol]), chart_type: 'column' });
                    } else if (isLineRequest) {
                      await createChartWithTwoColumns(catCol, numCol, 'line', `${headers[numCol]} Trend`);
                      createdChartMappings.push({ x_header: String(headers[catCol]), y_header: String(headers[numCol]), chart_type: 'line' });
                    } else {
                      await createChartWithTwoColumns(catCol, numCol, 'column', `${headers[numCol]} by ${headers[catCol]}`);
                      createdChartMappings.push({ x_header: String(headers[catCol]), y_header: String(headers[numCol]), chart_type: 'column' });
                    }
                  } else if (chartsCreated === 0 && isPieRequest) {
                    if (categoricalCols.length > 0 && numericCols.length > 0) {
                      await createChartWithTwoColumns(categoricalCols[0], numericCols[0], 'pie', `${headers[categoricalCols[0]]} Distribution`);
                    } else if (categoricalCols.length > 0 && numCols >= 2) {
                      await createChartWithTwoColumns(categoricalCols[0], 0, 'pie', `${headers[categoricalCols[0]]} Count`);
                    } else if (numCols >= 2) {
                      await createChartWithTwoColumns(0, 1, 'pie', `${headers[0]} Distribution`);
                    } else {
                      await createChartForColumn(0, 'pie', `${headers[0]} Distribution`);
                    }
                  } else if (chartsCreated === 0 && isBarRequest) {
                    if (numericCols.length > 0 && categoricalCols.length > 0) {
                      await createChartWithTwoColumns(categoricalCols[0], numericCols[0], 'column', `${headers[numericCols[0]]} by ${headers[categoricalCols[0]]}`);
                    } else if (numericCols.length > 0) {
                      await createChartForColumn(numericCols[0], 'column', `${headers[numericCols[0]]} Distribution`);
                    } else if (numCols >= 2) {
                      await createChartWithTwoColumns(0, 1, 'column', `${headers[1]} Distribution`);
                    } else {
                      await createChartForColumn(0, 'column', `${headers[0]} Distribution`);
                    }
                  } else if (chartsCreated === 0 && isLineRequest) {
                    if (numericCols.length > 0 && categoricalCols.length > 0) {
                      await createChartWithTwoColumns(categoricalCols[0], numericCols[0], 'line', `${headers[numericCols[0]]} Trend`);
                    } else if (numericCols.length > 0 && numCols >= 2) {
                      await createChartWithTwoColumns(0, numericCols[0], 'line', `${headers[numericCols[0]]} Trend`);
                    } else if (numericCols.length > 0) {
                      await createChartForColumn(numericCols[0], 'line', `${headers[numericCols[0]]} Trend`);
                    } else if (numCols >= 2) {
                      await createChartWithTwoColumns(0, 1, 'line', `${headers[1]} Trend`);
                    } else {
                      await createChartForColumn(0, 'line', `${headers[0]} Trend`);
                    }
                  } else if (chartsCreated === 0) {
                    if (numericCols.length > 0 && categoricalCols.length > 0) {
                      await createChartWithTwoColumns(categoricalCols[0], numericCols[0], 'column', `${headers[numericCols[0]]} by ${headers[categoricalCols[0]]}`);
                    } else if (numericCols.length > 0) {
                      await createChartForColumn(numericCols[0], 'line', `${headers[numericCols[0]]} Trend`);
                    } else if (categoricalCols.length > 0) {
                      await createChartForColumn(categoricalCols[0], 'column', `${headers[categoricalCols[0]]} Count`);
                    } else if (numCols >= 2) {
                      await createChartWithTwoColumns(0, 1, 'column', `${headers[1]} Distribution`);
                    } else {
                      await createChartForColumn(0, 'column', 'Data Chart');
                    }
                  }

                  if (chartsCreated === 0) {
                    await ExcelAPI.createChart('column', address, sheetName, 'Data Chart');
                    chartsCreated++;
                    chartDesc = `📊 Chart created on ${sheetName}`;
                  }

                  const expectedMappings = backendChartRequests.length > 0
                    ? backendChartRequests.map((r) => ({
                        x_header: String(r.x_header || headers[r.x_col_index] || ''),
                        y_header: String(r.y_header || headers[r.y_col_index] || ''),
                        chart_type: r.chart_type || 'column',
                      }))
                    : createdChartMappings;
                  let verificationNote = '';
                  if (expectedMappings.length > 0 && createdChartMappings.length > 0) {
                    setProcessingPhase('verifying');
                    try {
                      const verifyRes = await fetch('/api/verify-chart-execution', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                          request_message: userMessage,
                          expected_charts: expectedMappings,
                          created_charts: createdChartMappings,
                        }),
                      });
                      if (verifyRes.ok) {
                        const verify = await verifyRes.json();
                        verificationNote = verify.ok
                          ? `\n✅ Verification: ${verify.message}`
                          : `\n⚠️ Verification: ${verify.message}`;
                      }
                    } catch {}
                  }

                  const aiMsg: ExtendedMessage = {
                    id: `msg-${Date.now()}`,
                    role: 'assistant',
                    content: `${analysisResult.analysis || 'Analysis complete.'}\n\n${chartDesc}${verificationNote}`,
                  };
                  setMessages((prev) => [...prev, aiMsg]);
                  setIsProcessing(false);
                  setProcessingPhase('idle');
                  return;
                } catch (chartErr: any) {
                  const aiMsg: ExtendedMessage = {
                    id: `msg-${Date.now()}`,
                    role: 'assistant',
                    content: `${analysisResult.analysis || 'Analysis complete.'}\n\n⚠️ Chart creation failed: ${chartErr.message}`,
                  };
                  setMessages((prev) => [...prev, aiMsg]);
                  setIsProcessing(false);
                  setProcessingPhase('idle');
                  return;
                }
              }

              if (isAnalysisRequest) {
              const aiMsg: ExtendedMessage = {
                id: `msg-${Date.now()}`,
                role: 'assistant',
                content: analysisResult.analysis || 'Analysis complete.',
              };
              setMessages((prev) => [...prev, aiMsg]);
              setIsProcessing(false);
              setProcessingPhase('idle');
              return;
            }
          }
        }

        // Guardrails: Check for destructive operations
        const hasDestructiveAction = plan.some((s: any) => 
          ['delete_worksheet', 'delete_rows', 'delete_columns'].includes(s.action)
        );
        if (hasDestructiveAction) {
          const confirmDelete = await confirm(
            'This will delete data. Are you sure you want to continue?',
            { title: 'Confirm Delete', confirmVariant: 'danger', confirmLabel: 'Delete' }
          );
          if (!confirmDelete) {
            setMessages((prev) => [...prev, {
              id: `ai-${Date.now()}`,
              role: 'assistant',
              content: 'Operation cancelled. No changes were made.',
            }]);
            setIsProcessing(false);
            setProcessingPhase('idle');
            return;
          }
        }

        // Guardrails: Warn on large write operations
        const largeWriteStep = plan.find((s: any) => 
          s.action === 'set_values' && s.params?.values?.length > 1000
        );
        if (largeWriteStep) {
          const confirmLarge = await confirm(
            `This will write ${largeWriteStep.params.values.length} rows. Continue?`,
            { title: 'Large Data Write', confirmVariant: 'warning' }
          );
          if (!confirmLarge) {
            setMessages((prev) => [...prev, {
              id: `ai-${Date.now()}`,
              role: 'assistant',
              content: 'Operation cancelled. Large write prevented.',
            }]);
            setIsProcessing(false);
            setProcessingPhase('idle');
            return;
          }
        }

        // Guardrails: Warn if writing to different address than selected
        if (userWantsSelected && selectedRange) {
          const writeStep = plan.find((s: any) => 
            s.action === 'set_values' && s.params?.address
          );
          if (writeStep && writeStep.params.address !== selectedRange.address) {
            setProcessingPhase('confirming');
            const confirmAddress = await confirm(
              `You selected ${selectedRange.address} but the plan wants to write to ${writeStep.params.address}. Use your selection (${selectedRange.address}) instead?`,
              { title: 'Address Mismatch', confirmVariant: 'warning' }
            );
            if (!confirmAddress) {
              setProcessingPhase('idle');
              return;
            }
          }
        }

        // Pre-execution Review: Show plan in chat (no modal confirmation for scale)
        const planSummary = plan.map((s: any, idx: number) => 
          `${idx + 1}. ${s.action}${s.params?.address ? ` (${s.params.address})` : ''}: ${s.description || ''}`
        ).join('\n');
        validationErrors.push(`Plan: ${planSummary}`);

        // Error Detection: Check for potential Excel errors before execution
        const potentialErrors: string[] = [];
        for (const step of plan) {
          if (step.action === 'set_formulas' && step.params?.formulas) {
            for (const row of step.params.formulas) {
              if (Array.isArray(row)) {
                for (const cell of row) {
                  if (typeof cell === 'string' && cell.toUpperCase().includes('VLOOKUP')) {
                    potentialErrors.push(`VLOOKUP in ${step.params.address} may return #N/A if lookup value not found - consider wrapping in IFERROR`);
                  }
                  if (typeof cell === 'string' && cell.toUpperCase().includes('/')) {
                    const parts = cell.split('/');
                    if (parts.length > 1 && parts[1].match(/[A-Z]+\d+|SUM\(|COUNT\(/i)) {
                      potentialErrors.push(`Division in ${step.params.address} may return #DIV/0! if denominator is zero - consider IFERROR`);
                    }
                  }
                }
              }
            }
          }
          if (step.action === 'set_values' && step.params?.values) {
            let emptyCells = 0;
            for (const row of step.params.values) {
              if (Array.isArray(row)) {
                for (const cell of row) {
                  if (cell === '' || cell === null || cell === undefined) emptyCells++;
                }
              }
            }
            if (emptyCells > 0 && step.params.values.length > 1) {
              potentialErrors.push(`Data in ${step.params.address} has ${emptyCells} empty cells - may cause calculation errors`);
            }
            for (let col = 0; col < (step.params.values[0]?.length || 0); col++) {
              let numCount = 0, textCount = 0;
              for (let row = 1; row < Math.min(step.params.values.length, 10); row++) {
                const cell = step.params.values[row]?.[col];
                if (typeof cell === 'number') numCount++;
                else if (typeof cell === 'string' && cell !== '') textCount++;
              }
              if (numCount > 0 && textCount > 0) {
                potentialErrors.push(`Column ${col + 1} in ${step.params.address} has mixed text/numbers - may cause calculation issues`);
              }
            }
          }
        }
        if (potentialErrors.length > 0) {
          const userConfirmed = await confirm(
            `Potential issues detected:\n${potentialErrors.slice(0, 3).join('\n')}\n\nContinue anyway?`,
            { title: 'Error Detection Warning', confirmVariant: 'warning', confirmLabel: 'Continue' }
          );
          if (!userConfirmed) {
            setMessages((prev) => [...prev, {
              id: `ai-${Date.now()}`,
              role: 'assistant',
              content: 'Operation cancelled due to potential errors.',
            }]);
            setIsProcessing(false);
            setProcessingPhase('idle');
            return;
          }
          validationErrors.push(...potentialErrors.map(e => `Warning: ${e}`));
        }

        setProcessingPhase('executing');
        const executionResults: { action: string; success: boolean; output: string }[] = [];
        let lastCreatedSheet: string | null = null;
        let lastDeletedSheet: string | null = null;
        const knownSheets = new Set<string>();
        let completedSteps = 0;
        const totalSteps = validSteps.length;
        
        // Operation Timeout: Max 2 minutes for execution
        const OPERATION_TIMEOUT_MS = 120000;
        const STEP_TIMEOUT_MS = 30000;
        const executionStartTime = Date.now();
        
        // Check for timeout before each step
        const checkTimeout = () => {
          const elapsed = Date.now() - executionStartTime;
          if (elapsed > OPERATION_TIMEOUT_MS) {
            throw new Error(`Operation timed out after ${Math.round(elapsed/1000)}s. Please try again with fewer steps.`);
          }
        };

        if (workbookContext) {
          try {
            const ctx = JSON.parse(workbookContext);
            ctx.sheets?.forEach((s: string) => knownSheets.add(s));
          } catch {}
        }

        // Add diagnostic info to execution results
        const executionSummary = `[DEBUG] Selected range detected: ${selectedRange?.address || 'NONE'} in sheet ${selectedRange?.sheetName || 'NONE'}. User wants selected: ${userWantsSelected}`;
        console.log(executionSummary);
        
        for (let i = 0; i < validSteps.length; i++) {
          // Operation Timeout: Check before each step
          checkTimeout();
          
          const step = validSteps[i];
          const params = { ...step.params };

          // Update live step progress
          setExecutionProgress({ current: i + 1, total: validSteps.length, action: step.action });

          // CRITICAL: Force selected range for write operations when user refers to selection
          console.log('[EXECUTION] Final check - userWantsSelected:', userWantsSelected, '| react selectedRange state:', selectedRange, '| step.action:', step.action, '| original params:', { address: params.address, sheet_name: params.sheet_name });
          if (userWantsSelected && selectedRange) {
            const writeActions = ['set_values', 'set_formulas', 'apply_format', 'create_chart', 'create_table', 'sort_range', 'auto_fill'];
            if (writeActions.includes(step.action)) {
              console.log('[EXECUTION] Overwriting params with selected range:', selectedRange.address, selectedRange.sheetName);
              // ALWAYS force to selected sheet and address for "put here" type requests
              params.address = selectedRange.address;
              params.sheet_name = selectedRange.sheetName;
              validationErrors.push(`Step ${i + 1}: Using selected range "${selectedRange.address}" in "${selectedRange.sheetName}"`);
            }
            // For charts, use selected range as data source
            if (step.action === 'create_chart' && !params.data_range) {
              params.data_range = selectedRange.address;
              params.sheet_name = selectedRange.sheetName;
              validationErrors.push(`Step ${i + 1}: Using selected range as chart data source`);
            }
          }

          if (step.action === 'add_worksheet') {
            if (!params.name || typeof params.name !== 'string' || params.name.trim() === '') {
              params.name = `Sheet_${Date.now()}`;
              validationErrors.push(`Step ${i + 1}: Missing sheet name — auto-generated "${params.name}"`);
            }
            params.name = params.name.replace(/[\/?*\[\]:]/g, '_').trim().slice(0, 31);
            if (!params.name || params.name.length === 0) {
              params.name = `Sheet_${Date.now()}`;
              validationErrors.push(`Step ${i + 1}: Invalid sheet name after sanitization — auto-generated "${params.name}"`);
            }
            lastCreatedSheet = params.name;
            lastDeletedSheet = null;
            knownSheets.add(params.name);
          }

          if (step.action === 'delete_worksheet') {
            const sheetToDelete = params.sheet_name || params.name;
            if (!sheetToDelete) {
              executionResults.push({ action: step.action, success: false, output: 'Skipped: no sheet name provided' });
              continue;
            }
            lastDeletedSheet = sheetToDelete;
            knownSheets.delete(sheetToDelete);
          }

          if (lastCreatedSheet && !lastDeletedSheet &&
              step.action !== 'add_worksheet' && step.action !== 'get_workbook_structure' && step.action !== 'get_selected_range') {
            if (!params.sheet_name || params.sheet_name === 'Sheet1' || params.sheet_name === 'Sheet') {
              params.sheet_name = lastCreatedSheet;
              validationErrors.push(`Step ${i + 1}: Using newly created sheet "${lastCreatedSheet}"`);
            }
          }

          const A1_REGEX = /^[A-Z]{1,3}\d+(:[A-Z]{1,3}\d+)?$/;
          const addressFields = ['address', 'source_address', 'target_address', 'data_range'];
          for (const field of addressFields) {
            if (params[field] && typeof params[field] === 'string') {
              const clean = params[field].includes('!') ? params[field].split('!')[1] : params[field];
              if (!A1_REGEX.test(clean)) {
                const fallback = field === 'data_range' ? 'A1:B10' : field.includes('source') ? 'A1' : field.includes('target') ? 'A1:A10' : 'A1';
                validationErrors.push(`Step ${i + 1}: Invalid range "${params[field]}" for ${field} — using "${fallback}"`);
                params[field] = fallback;
              }
            }
          }

          if (step.action === 'create_chart') {
            const validTypes = ['column', 'bar', 'line', 'pie', 'pie3D', 'doughnut', 'area', 'scatter', 'radar', 'surface', 'bubble'];
            if (!params.chart_type || !validTypes.includes(params.chart_type)) {
              params.chart_type = 'column';
              validationErrors.push(`Step ${i + 1}: Invalid chart_type — defaulting to "column"`);
            }
            if (!params.data_range) {
              params.data_range = 'A1:B10';
              validationErrors.push(`Step ${i + 1}: Missing data_range — using "A1:B10"`);
            }
            if (!params.sheet_name && lastCreatedSheet) {
              params.sheet_name = lastCreatedSheet;
              validationErrors.push(`Step ${i + 1}: Using newly created sheet "${lastCreatedSheet}" for chart`);
            }
            if (params.position) {
              const pos = params.position;
              if (typeof pos.left !== 'number' || pos.left < 0) pos.left = 300;
              if (typeof pos.top !== 'number' || pos.top < 0) pos.top = 50;
              if (typeof pos.width !== 'number' || pos.width < 50 || pos.width > 2000) pos.width = 600;
              if (typeof pos.height !== 'number' || pos.height < 50 || pos.height > 2000) pos.height = 400;
            }
          }

          if (step.action === 'set_values') {
            if (!params.values || !Array.isArray(params.values) || params.values.length === 0) {
              executionResults.push({ action: step.action, success: false, output: 'Skipped: empty or invalid values array' });
              continue;
            }
            if (params.values[0] && (!Array.isArray(params.values[0]) || params.values[0].length === 0)) {
              executionResults.push({ action: step.action, success: false, output: 'Skipped: values array has empty rows' });
              continue;
            }
            if (params.values.length > 5000) {
              params.values = params.values.slice(0, 5000);
              validationErrors.push(`Step ${i + 1}: Dataset too large — clamped to 5000 rows`);
            }
            
            // ONLY auto-calculate address if user did NOT want selected range
            // If userWantsSelected is true, preserve the selected range address
            if (!userWantsSelected || !selectedRange) {
              const rows = params.values.length;
              const cols = params.values[0]?.length || 1;
              const colLetter = cols <= 26 ? String.fromCharCode(64 + cols) : String.fromCharCode(64 + Math.floor((cols - 1) / 26)) + String.fromCharCode(64 + ((cols - 1) % 26) + 1);
              const correctAddress = `A1:${colLetter}${rows}`;
              
              if (params.address && params.address !== correctAddress) {
                validationErrors.push(`Step ${i + 1}: Address "${params.address}" doesn't match data size — corrected to "${correctAddress}"`);
              }
              params.address = correctAddress;
            } else {
              // User wants selected range - validate/adjust data to match range size
              const selectedRows = selectedRange.rowCount;
              const selectedCols = selectedRange.columnCount;
              const dataRows = params.values.length;
              const dataCols = params.values[0]?.length || 1;
              
              // If data dimensions don't match selected range, adjust data array
              if (dataRows !== selectedRows || dataCols !== selectedCols) {
                validationErrors.push(`Step ${i + 1}: Data size ${dataRows}×${dataCols} doesn't match selected range ${selectedRows}×${selectedCols} — adjusting...`);
                
                // Pad or trim to match selected range
                const adjustedValues: (string | number | boolean | null)[][] = [];
                for (let r = 0; r < selectedRows; r++) {
                  const row: (string | number | boolean | null)[] = [];
                  for (let c = 0; c < selectedCols; c++) {
                    if (r < dataRows && c < dataCols) {
                      row.push(params.values[r][c]);
                    } else {
                      row.push(null);
                    }
                  }
                  adjustedValues.push(row);
                }
                params.values = adjustedValues;
                validationErrors.push(`Step ${i + 1}: Data adjusted to ${selectedRows}×${selectedCols} for selected range`);
              }
              // Use selected range address and sheet
              params.address = selectedRange.address;
              params.sheet_name = selectedRange.sheetName;
              validationErrors.push(`Step ${i + 1}: Using selected range "${selectedRange.address}" in "${selectedRange.sheetName}"`);
            }
            
            // Only use lastCreatedSheet if user did NOT want selected range
            if ((!userWantsSelected || !selectedRange) && !params.sheet_name && lastCreatedSheet) {
              params.sheet_name = lastCreatedSheet;
              validationErrors.push(`Step ${i + 1}: Using sheet "${lastCreatedSheet}" for data`);
            }
          }

          if (step.action === 'set_formulas') {
            if (!params.formulas || !Array.isArray(params.formulas) || params.formulas.length === 0) {
              executionResults.push({ action: step.action, success: false, output: 'Skipped: empty or invalid formulas array' });
              continue;
            }
            for (const row of params.formulas) {
              if (Array.isArray(row)) {
                for (let j = 0; j < row.length; j++) {
                  if (typeof row[j] === 'string' && row[j].length > 0 && !row[j].startsWith('=') && !row[j].startsWith('"')) {
                    row[j] = `=${row[j]}`;
                  }
                }
              }
            }
            if (!params.address) {
              params.address = 'A1';
              validationErrors.push(`Step ${i + 1}: Missing address — using "A1"`);
            }
            // Formula validation - check for common errors
            let formulaErrors: string[] = [];
            for (const row of params.formulas) {
              if (Array.isArray(row)) {
                for (const cell of row) {
                  if (typeof cell === 'string' && cell.startsWith('=')) {
                    const f = cell.toUpperCase();
                    // Check for circular reference patterns
                    if (f.match(/=[A-Z]+\d+:\s*[A-Z]+\d+/)) {
                      formulaErrors.push(`Potential range error in: ${cell.slice(0, 20)}...`);
                    }
                    // Check for unclosed parentheses
                    const openParens = (f.match(/\(/g) || []).length;
                    const closeParens = (f.match(/\)/g) || []).length;
                    if (openParens !== closeParens) {
                      formulaErrors.push(`Unbalanced parentheses: ${cell.slice(0, 20)}...`);
                    }
                    // Check for missing arguments in common functions
                    if (/\bVLOOKUP\s*\(\s*,/i.test(f) || /\bIF\s*\(\s*,/i.test(f)) {
                      formulaErrors.push(`Missing argument: ${cell.slice(0, 20)}...`);
                    }
                  }
                }
              }
            }
            if (formulaErrors.length > 0) {
              validationErrors.push(`Formula warnings: ${formulaErrors.slice(0, 3).join('; ')}`);
            }
          }

          if (step.action === 'apply_format') {
            const hasBorderParams = params.border_all || params.border_color || params.border_style || params.border_weight;
            if (!hasBorderParams && /border/i.test(userMessage)) {
              const msg = userMessage.toLowerCase();
              const borderColorMatch = msg.match(/\b(red|blue|green|black|white|grey|gray|yellow|orange|purple|pink|brown|cyan|magenta|navy|teal|maroon|olive|gold|silver|copper|bronze)\b/i);
              const colorToHex: Record<string, string> = {
                red: '#FF0000', blue: '#0000FF', green: '#008000', black: '#000000', white: '#FFFFFF',
                grey: '#808080', gray: '#808080', yellow: '#FFFF00', orange: '#FFA500', purple: '#800080',
                pink: '#FFC0CB', brown: '#A52A2A', cyan: '#00FFFF', magenta: '#FF00FF', navy: '#000080',
                teal: '#008080', maroon: '#800000', olive: '#808000', gold: '#FFD700', silver: '#C0C0C0',
                copper: '#B87333', bronze: '#CD7F32'
              };
              params.border_all = true;
              params.border_color = borderColorMatch ? (colorToHex[borderColorMatch[1].toLowerCase()] || '#000000') : '#000000';
              if (/thick/i.test(msg)) params.border_weight = 'Thick';
              else if (/medium/i.test(msg)) params.border_weight = 'Medium';
              else if (/hairline/i.test(msg)) params.border_weight = 'Hairline';
              else params.border_weight = 'Thin';
              if (/dash/i.test(msg)) params.border_style = 'Dash';
              else if (/dot/i.test(msg)) params.border_style = 'Dot';
              else if (/double/i.test(msg)) params.border_style = 'Double';
              else params.border_style = 'Continuous';
              validationErrors.push(`Step ${i + 1}: Auto-injected border (color: ${params.border_color}, style: ${params.border_style}, weight: ${params.border_weight})`);
            }
            const HEX_REGEX = /^#[0-9A-Fa-f]{6}$/;
            const COLOR_MAP: Record<string, string> = {
              grey: '#808080', gray: '#808080', lightgrey: '#D3D3D3', lightgray: '#D3D3D3',
              darkgrey: '#A9A9A9', darkgray: '#A9A9A9', silver: '#C0C0C0',
              red: '#FF0000', blue: '#0000FF', green: '#008000', yellow: '#FFFF00',
              white: '#FFFFFF', black: '#000000', orange: '#FFA500', purple: '#800080',
              pink: '#FFC0CB', brown: '#A52A2A', cyan: '#00FFFF', magenta: '#FF00FF',
              navy: '#000080', teal: '#008080', maroon: '#800000', olive: '#808000',
              lime: '#00FF00', aqua: '#00FFFF', fuchsia: '#FF00FF', gold: '#FFD700',
              coral: '#FF7F50', salmon: '#FA8072', tomato: '#FF6347', indigo: '#4B0082',
              violet: '#EE82EE', plum: '#DDA0DD', lavender: '#E6E6FA', beige: '#F5F5DC',
              wheat: '#F5DEB3', tan: '#D2B48C', khaki: '#F0E68C', mint: '#98FF98',
              mintcream: '#F5FFFA', honeydew: '#F0FFF0', azure: '#F0FFFF', ivory: '#FFFFF0',
              crimson: '#DC143C', scarlet: '#FF2400', ruby: '#E0115F', emerald: '#50C878',
              jade: '#00A86B', amber: '#FFBF00', bronze: '#CD7F32', copper: '#B87333',
              silver2: '#C0C0C0', steel: '#4682B4', slate: '#708090', charcoal: '#36454F',
              skyblue: '#87CEEB', royalblue: '#4169E1', dodgerblue: '#1E90FF', deepskyblue: '#00BFFF',
              lightblue: '#ADD8E6', powderblue: '#B0E0E6', paleturquoise: '#AFEEEE', lightcyan: '#E0FFFF',
            };
            if (params.font_color) {
              const fc = params.font_color.toLowerCase().trim();
              if (COLOR_MAP[fc]) {
                params.font_color = COLOR_MAP[fc];
              } else if (!HEX_REGEX.test(params.font_color)) {
                delete params.font_color;
                validationErrors.push(`Step ${i + 1}: Invalid font_color — removed`);
              }
            }
            if (params.fill_color) {
              const fl = params.fill_color.toLowerCase().trim();
              if (COLOR_MAP[fl]) {
                params.fill_color = COLOR_MAP[fl];
              } else if (!HEX_REGEX.test(params.fill_color)) {
                delete params.fill_color;
                validationErrors.push(`Step ${i + 1}: Invalid fill_color — removed`);
              }
            }
            if (params.font_size && (typeof params.font_size !== 'number' || params.font_size < 1 || params.font_size > 400)) {
              params.font_size = Math.max(1, Math.min(400, Number(params.font_size) || 11));
              validationErrors.push(`Step ${i + 1}: Font size out of range — clamped to ${params.font_size}`);
            }
            const validHAlign = ['Left', 'Center', 'Right', 'Justify', 'Distributed'];
            const validVAlign = ['Top', 'Center', 'Bottom', 'Justify', 'Distributed'];
            if (params.horizontal_alignment && !validHAlign.includes(params.horizontal_alignment)) {
              delete params.horizontal_alignment;
              validationErrors.push(`Step ${i + 1}: Invalid horizontal_alignment — removed`);
            }
            if (params.vertical_alignment && !validVAlign.includes(params.vertical_alignment)) {
              delete params.vertical_alignment;
              validationErrors.push(`Step ${i + 1}: Invalid vertical_alignment — removed`);
            }
            if (!params.address) {
              params.address = 'A1';
              validationErrors.push(`Step ${i + 1}: Missing address — using "A1"`);
            }
          }

          if (['insert_rows', 'delete_rows', 'insert_columns', 'delete_columns'].includes(step.action)) {
            if (typeof params.count !== 'number' || params.count < 1) {
              params.count = 1;
              validationErrors.push(`Step ${i + 1}: Invalid count — using 1`);
            }
            if (params.count > 1000) {
              params.count = 1000;
              validationErrors.push(`Step ${i + 1}: Count too large — clamped to 1000`);
            }
            if (!params.address) {
              params.address = 'A1';
              validationErrors.push(`Step ${i + 1}: Missing address — using "A1"`);
            }
          }

          if (step.action === 'sort_range') {
            if (typeof params.column_index !== 'number' || params.column_index < 0) {
              params.column_index = 0;
              validationErrors.push(`Step ${i + 1}: Invalid column_index — using 0`);
            }
            if (params.ascending === undefined) params.ascending = true;
            if (!params.address) {
              params.address = 'A1';
              validationErrors.push(`Step ${i + 1}: Missing address — using "A1"`);
            }
          }

          if (step.action === 'auto_fill') {
            if (!params.source_address) params.source_address = 'A1';
            if (!params.target_address) params.target_address = 'A1:A10';
          }

          if (step.action === 'create_table') {
            if (!params.name || typeof params.name !== 'string') {
              params.name = `Table_${Date.now()}`;
              validationErrors.push(`Step ${i + 1}: Missing table name — auto-generated`);
            }
            params.name = params.name.replace(/[^a-zA-Z0-9_]/g, '_');
            
            // Use selected range if user wants selected, otherwise use AI-provided address
            if (userWantsSelected && selectedRange) {
              params.address = selectedRange.address;
              params.sheet_name = selectedRange.sheetName;
              validationErrors.push(`Step ${i + 1}: Creating table in selected range "${selectedRange.address}"`);
            } else {
              if (!params.address) {
                params.address = 'A1';
                validationErrors.push(`Step ${i + 1}: Missing address — using "A1"`);
              }
              if (!params.sheet_name && lastCreatedSheet) {
                params.sheet_name = lastCreatedSheet;
                validationErrors.push(`Step ${i + 1}: Using sheet "${lastCreatedSheet}" for table`);
              }
            }
          }

          if (step.action === 'add_worksheet' && knownSheets.has(params.name)) {
            params.name = `${params.name}_${Date.now()}`;
            validationErrors.push(`Step ${i + 1}: Sheet already exists — renamed to "${params.name}"`);
            knownSheets.add(params.name);
          }

          if (params.sheet_name && params.sheet_name === lastDeletedSheet) {
            executionResults.push({ action: step.action, success: false, output: `Skipped: references deleted sheet "${params.sheet_name}"` });
            continue;
          }

          for (const field of addressFields) {
            if (params[field] && typeof params[field] === 'string' && params[field].includes('!')) {
              const refSheet = params[field].split('!')[0].replace(/'/g, '');
              if (refSheet && refSheet !== params.sheet_name && !knownSheets.has(refSheet)) {
                validationErrors.push(`Step ${i + 1}: Cross-sheet reference to unknown sheet "${refSheet}" — using active sheet`);
                params[field] = params[field].split('!')[1] || params[field];
              }
            }
          }

          const MAX_ROW = 1048576;
          const MAX_COL = 16384;
          for (const field of addressFields) {
            if (params[field] && typeof params[field] === 'string' && A1_REGEX.test(params[field])) {
              const parts = params[field].split(':');
              for (const part of parts) {
                const colMatch = part.match(/^([A-Z]+)/);
                const rowMatch = part.match(/(\d+)$/);
                if (rowMatch) {
                  const rowNum = parseInt(rowMatch[1], 10);
                  if (rowNum > MAX_ROW) {
                    validationErrors.push(`Step ${i + 1}: Row ${rowNum} exceeds Excel max (${MAX_ROW}) — clamped`);
                    params[field] = params[field].replace(rowMatch[1], String(MAX_ROW));
                  }
                  if (rowNum < 1) {
                    validationErrors.push(`Step ${i + 1}: Row ${rowNum} invalid — clamped to 1`);
                    params[field] = params[field].replace(rowMatch[1], '1');
                  }
                }
                if (colMatch) {
                  const colStr = colMatch[1];
                  let colNum = 0;
                  for (let c = 0; c < colStr.length; c++) {
                    colNum = colNum * 26 + (colStr.charCodeAt(c) - 64);
                  }
                  if (colNum > MAX_COL) {
                    validationErrors.push(`Step ${i + 1}: Column ${colStr} exceeds Excel max (XFD) — clamped`);
                    params[field] = params[field].replace(colStr, 'XFD');
                  }
                }
              }
            }
          }

          if (step.action === 'set_values' && params.address && params.values) {
            const addrMatch = params.address.match(/^([A-Z]+)(\d+)/);
            if (addrMatch) {
              const startCol = addrMatch[1];
              const startRow = parseInt(addrMatch[2], 10);
              let colSpan = 0;
              for (let c = 0; c < startCol.length; c++) {
                colSpan = colSpan * 26 + (startCol.charCodeAt(c) - 64);
              }
              const rowSpan = params.values.length;
              const colCount = params.values[0]?.length || 0;
              const endRow = startRow + rowSpan - 1;
              const endColNum = colSpan + colCount - 1;
              if (endRow > MAX_ROW) {
                params.values = params.values.slice(0, MAX_ROW - startRow + 1);
                validationErrors.push(`Step ${i + 1}: Values exceed row limit — truncated`);
              }
              if (endColNum > MAX_COL) {
                params.values = params.values.map((row: any[]) => row.slice(0, MAX_COL - colSpan + 1));
                validationErrors.push(`Step ${i + 1}: Values exceed column limit — truncated`);
              }
            }
          }

          if (step.action === 'set_values' && params.values) {
            const sanitizeValue = (v: any): any => {
              if (typeof v === 'number' && (Number.isNaN(v) || !Number.isFinite(v))) return 0;
              if (typeof v === 'string' && (v === 'NaN' || v === 'Infinity' || v === '-Infinity')) return 0;
              if (v === null || v === undefined) return '';
              return v;
            };
            params.values = params.values.map((row: any[]) => row.map((cell: any) => sanitizeValue(cell)));
          }

          if (step.action === 'set_values' && params.values) {
            params.values = params.values.map((row: any[]) =>
              row.map((cell: any) => {
                if (typeof cell === 'string') {
                  if (cell.startsWith('=') || cell.startsWith('+') || cell.startsWith('-') || cell.startsWith('@')) {
                    return `'${cell}`;
                  }
                  return cell;
                }
                return cell;
              })
            );
          }

          if (validSteps.length === 0) {
            executionResults.push({ action: 'plan', success: false, output: 'No valid steps in plan — AI returned empty or invalid plan' });
          }

          const fixedStep = { ...step, params };

          try {
            if (typeof Excel === 'undefined') {
              executionResults.push({ action: step.action, success: false, output: 'Excel is not ready. Please refresh the add-in.' });
              continue;
            }

            if (step.action === 'set_values' && lastCreatedSheet) {
              await new Promise(r => setTimeout(r, 1000));
              
              const freshContext = await getWorkbookContext();
              if (freshContext) {
                const ctx = JSON.parse(freshContext);
                const newSheet = ctx.sheets.find((s: string) => s.includes(lastCreatedSheet!.substring(0, 10)));
                if (newSheet) {
                  params.sheet_name = newSheet;
                  validationErrors.push(`Step ${i + 1}: Using verified sheet "${newSheet}"`);
                }
              }
            }
            
            if (step.action === 'create_table' && lastCreatedSheet) {
              params.sheet_name = lastCreatedSheet;
              validationErrors.push(`Step ${i + 1}: Force using sheet "${lastCreatedSheet}" for table`);
            }
            
            let retryCount = 0;
            let lastError = '';
            let lastSuccess = false;
            let successOutput = '';

            // Checkpoint: Snapshot cells before mutation
            const writeActions = ['set_values', 'set_formulas', 'apply_format', 'delete_rows', 'delete_columns', 'insert_rows', 'insert_columns'];
            if (writeActions.includes(step.action) && params.address) {
              try {
                let sheetName = params.sheet_name;
                if (!sheetName) {
                  const ctx = await getWorkbookContext();
                  if (ctx) {
                    sheetName = JSON.parse(ctx).activeSheet;
                  }
                }
                if (sheetName) {
                  const currentData = await ExcelAPI.getRange(params.address, sheetName);
                  if (currentData && currentData.values) {
                    addCheckpoint({
                      sheet: sheetName,
                      address: params.address,
                      values: currentData.values,
                      description: `Before ${step.action}: ${step.description || step.action}`,
                    });
                  }
                }
              } catch {}
            }

            while (retryCount < 2) {
              try {
                const result = await executeStep(fixedStep);
                successOutput = result;
                lastSuccess = true;
                break;
              } catch (error) {
                lastError = error instanceof Error ? error.message : String(error);
                
                if (step.action === 'add_worksheet' && lastError.includes('already exists')) {
                  knownSheets.add(params.name);
                  successOutput = `Using existing sheet "${params.name}"`;
                  lastSuccess = true;
                  break;
                }
                
                if (retryCount === 1) {
                  break;
                }
                retryCount++;
              }
            }
            
            if (lastSuccess) {
              executionResults.push({ action: step.action, success: true, output: successOutput });
              completedSteps++;
              // Track operation for memory
              trackOperation(step.action);
              // Goal monitoring - track progress
              if (totalSteps > 1) {
                validationErrors.push(`Progress: ${completedSteps}/${totalSteps} steps completed`);
              }
              if (step.action === 'add_worksheet') {
                const actualName = successOutput.match(/"([^"]+)"/)?.[1] || params.name;
                lastCreatedSheet = actualName;
                knownSheets.add(actualName);
                validationErrors.push(`Step ${i + 1}: Updated lastCreatedSheet to "${actualName}"`);
              }
            } else {
              executionResults.push({ action: step.action, success: false, output: lastError });
              // Record mistake for future avoidance
              addMistake(`${step.action} failed: ${lastError.slice(0, 100)}`, `Retry with corrected parameters`);
              
              // If table overlap error, add helpful message
              if (step.action === 'create_table' && lastError.toLowerCase().includes('overlap')) {
                executionResults.push({ action: step.action, success: false, output: `Hint: A table already exists in "${params.address}". Delete the existing table first, then retry if needed.` });
              }

              // ADAPTIVE EXECUTOR: Check if failure is critical — halt remaining steps and re-plan
              const isCritical = lastError.match(/sheet.*not found|not found|invalid range|does not exist|already deleted|excel is not ready/i);
              const remainingSteps = validSteps.slice(i + 1);
              if (isCritical && remainingSteps.length > 0) {
                validationErrors.push(`Adaptive: Step ${i + 1} failed critically — halting ${remainingSteps.length} remaining step(s)`);
                
                // Trigger immediate reflection for remaining steps
                if (apiKey) {
                  try {
                    setProcessingPhase('validating');
                    const adaptiveReflect = await fetch('/api/reflect', {
                      method: 'POST',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({
                        message: userMessage,
                        plan: remainingSteps,
                        results: remainingSteps.map(s => ({ action: s.action, success: false, output: `Depends on failed step: ${step.action}` })),
                        provider: settings.provider,
                        model: effectiveModel,
                        api_key: apiKey,
                      }),
                    });
                    if (adaptiveReflect.ok) {
                      const adaptiveData = await adaptiveReflect.json();
                      const adaptiveRecovery = adaptiveData.recovery || [];
                      for (const rec of adaptiveRecovery) {
                        try {
                          const recResult = await executeStep({ action: rec.action, params: rec.params, description: rec.description });
                          executionResults.push({ action: rec.action, success: true, output: `Adaptive: ${recResult}` });
                        } catch (recErr) {
                          executionResults.push({ action: rec.action, success: false, output: `Adaptive recovery failed: ${recErr instanceof Error ? recErr.message : String(recErr)}` });
                        }
                      }
                    }
                  } catch {}
                }
                break; // Stop executing remaining steps
              }
            }
          } catch (error) {
            const errMsg = error instanceof Error ? error.message : String(error);
            executionResults.push({ action: step.action, success: false, output: errMsg });
          }
        }

        // Reflection: If any step failed, try to auto-recover
        const failedSteps = executionResults.filter((r) => !r.success);
        if (failedSteps.length > 0) {
          setProcessingPhase('validating');
          console.log('[REFLECTION] Reflection triggered -', failedSteps.length, 'step(s) failed, attempting recovery...');
          
          // Check if API key is available
          if (!apiKey) {
            console.log('[REFLECTION] No API key - skipping reflection');
            validationErrors.push('Reflection skipped: No API key configured');
          } else {
            try {
              const reflectionResponse = await fetch('/api/reflect', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                  message: userMessage,
                  plan: validSteps,
                  results: executionResults,
                  provider: settings.provider,
                  model: effectiveModel,
                  api_key: apiKey,
                }),
              });
              
              if (reflectionResponse.ok) {
                const reflectionData = await reflectionResponse.json();
                const recoverySteps = reflectionData.recovery || [];
                
                if (recoverySteps.length > 0) {
                  console.log('[REFLECTION] Recovery plan received -', recoverySteps.length, 'step(s)');
                  validationErrors.push(`Reflection: Attempting ${recoverySteps.length} recovery step(s)...`);
                  
                  let recoveryFailures = 0;
                  for (let j = 0; j < recoverySteps.length; j++) {
                    const recoveryStep = recoverySteps[j];
                    const recoveryParams = { ...recoveryStep.params };
                    
                    // Apply same selected range logic for recovery
                    if (userWantsSelected && selectedRange) {
                      const writeActions = ['set_values', 'set_formulas', 'apply_format', 'create_table'];
                      if (writeActions.includes(recoveryStep.action)) {
                        recoveryParams.address = selectedRange.address;
                        recoveryParams.sheet_name = selectedRange.sheetName;
                      }
                    }
                    
                    try {
                      const recoveryResult = await executeStep({ action: recoveryStep.action, params: recoveryParams, description: recoveryStep.description });
                      executionResults.push({ action: recoveryStep.action, success: true, output: `Recovery: ${recoveryResult}` });
                      validationErrors.push(`Recovery step ${j + 1}: ${recoveryStep.description} - SUCCESS`);
                    } catch (error) {
                      const errMsg = error instanceof Error ? error.message : String(error);
                      executionResults.push({ action: recoveryStep.action, success: false, output: `Recovery failed: ${errMsg}` });
                      validationErrors.push(`Recovery step ${j + 1}: ${recoveryStep.description} - FAILED`);
                      recoveryFailures++;
                      
                      // Skip remaining recovery if 2+ failures
                      if (recoveryFailures >= 2) {
                        validationErrors.push('Reflection: Too many recovery failures, stopping');
                        break;
                      }
                    }
                  }
                }
              } else {
                console.log('[REFLECTION] Reflection API returned error:', reflectionResponse.status);
              }
            } catch (e) {
              console.log('[REFLECTION] Reflection failed:', e);
            }
          }
        }
        
        setProcessingPhase('validating');
        const validationResponse = await fetch('/api/execute', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            message: userMessage,
            plan: plan,
            provider: settings.provider,
            model: effectiveModel,
            api_key: apiKey,
            results: executionResults,
          }),
        });

        let finalContent = responseContent;
        if (validationResponse.ok) {
          const validationData = await validationResponse.json();
          finalContent = validationData.content;
        }

        setMessages((prev) => [...prev, {
          id: `ai-${Date.now()}`,
          role: 'assistant',
          content: finalContent,
          planSteps: plan,
          executionResults,
        }]);
      } catch (error) {
        setMessages((prev) => [...prev, {
          id: `err-${Date.now()}`,
          role: 'assistant',
          content: `Error: ${error instanceof Error ? error.message : 'An unexpected error occurred'}. Make sure the backend is running.`,
        }]);
      } finally {
        setIsProcessing(false);
        setProcessingPhase('idle');
        setExecutionProgress(null);
      }
    },
    [settings, getApiKey]
  );

  const handleSend = useCallback((text: string, enableReasoning?: boolean) => {
    if (enableReasoning !== undefined) {
      setReasoningEnabled(enableReasoning);
    }
    processMessage(text);
    setTimeout(scrollToBottom, 100);
  }, [processMessage, scrollToBottom]);

  const handleToggleReasoning = useCallback((enabled: boolean) => {
    setReasoningEnabled(enabled);
  }, []);

  const handleSettingsSave = useCallback((newSettings: Settings) => {
    setSettings(newSettings);
    localStorage.setItem('excel-ai-settings', JSON.stringify(newSettings));
  }, []);

  const hasKey = getApiKey().length > 0;
  const phase = processingPhase ? PHASE_CONFIG[processingPhase] : PHASE_CONFIG.idle;

  return (
    <div className="flex flex-col" style={{ height: '100vh', maxHeight: '100vh', overflow: 'hidden' }}>
      {/* Premium Header */}
      <header className="flex-shrink-0 flex items-center justify-between px-3 py-2.5 border-b border-gray-200/60 bg-gradient-to-r from-[#217346] via-[#1E6B41] to-[#185C37]">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-lg bg-white/15 backdrop-blur-sm flex items-center justify-center ring-1 ring-white/20 shadow-sm">
            <svg className="w-4 h-4 text-white" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M12 2L2 7l10 5 10-5-10-5z" />
              <path d="M2 17l10 5 10-5" />
              <path d="M2 12l10 5 10-5" />
            </svg>
          </div>
          <div>
            <h1 className="text-[14px] font-black text-white tracking-tight leading-none" style={{ fontFamily: "'Inter', sans-serif" }}>Cel</h1>
            <p className="text-[9px] text-white/60 font-bold leading-none mt-0.5 tracking-widest uppercase">Agentic AI</p>
          </div>
        </div>
        <div className="flex items-center gap-1">
          <button
            onClick={handleUndo}
            className="w-7 h-7 rounded-lg hover:bg-white/10 flex items-center justify-center text-white/50 hover:text-white/80 transition-all"
            title="Undo last operation"
          >
            <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <polyline points="1 4 1 10 7 10" />
              <path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10" />
            </svg>
          </button>
          <button
            onClick={() => setMessages([])}
            className="w-7 h-7 rounded-lg hover:bg-white/10 flex items-center justify-center text-white/50 hover:text-white/80 transition-all"
            title="Clear chat"
          >
            <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M3 6h18M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6" />
            </svg>
          </button>
          <button
            onClick={() => setShowSettings(true)}
            className="w-7 h-7 rounded-lg hover:bg-white/10 flex items-center justify-center text-white/50 hover:text-white/80 transition-all"
            title="Settings"
          >
            <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="12" cy="12" r="3" />
              <path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42" />
            </svg>
          </button>
        </div>
      </header>

      {/* Processing Status Bar */}
      {isProcessing && (
        <div className="flex-shrink-0 px-3 py-1.5 bg-gradient-to-r from-violet-50/80 to-purple-50/80 border-b border-violet-100/80 flex items-center gap-2.5 animate-fade-in">
          <div className={cn('w-1.5 h-1.5 rounded-full animate-pulse', `bg-${phase.accent?.split(' ')[0]?.replace('from-', '') || 'violet-500'}`)} />
          <span className={cn('text-[11px] font-semibold', phase.color)}>
            {processingPhase === 'executing' && executionProgress
              ? `Step ${executionProgress.current}/${executionProgress.total}: ${executionProgress.action.replace(/_/g, ' ')}`
              : phase.label}
          </span>
          <div className="flex-1 h-1 bg-gray-200/60 rounded-full overflow-hidden">
            <div
              className={cn('h-full rounded-full transition-all duration-500')}
              style={{
                width:
                  processingPhase === 'understanding' ? '12%' :
                  processingPhase === 'context' ? '24%' :
                  processingPhase === 'schema' ? '36%' :
                  processingPhase === 'intent' ? '48%' :
                  processingPhase === 'planning' ? '60%' :
                  processingPhase === 'executing' ? '78%' :
                  processingPhase === 'verifying' || processingPhase === 'validating' ? '92%' :
                  processingPhase === 'thinking' || processingPhase === 'reasoning' || processingPhase === 'analyzing' ? '30%' :
                  '95%',
                background: `linear-gradient(90deg, ${
                  processingPhase === 'understanding' || processingPhase === 'thinking' || processingPhase === 'reasoning'
                    ? '#8b5cf6, #a78bfa'
                    : processingPhase === 'context'
                    ? '#6366f1, #60a5fa'
                    : processingPhase === 'schema' || processingPhase === 'analyzing'
                    ? '#06b6d4, #3b82f6'
                    : processingPhase === 'intent' || processingPhase === 'planning'
                    ? '#3b82f6, #60a5fa'
                    : processingPhase === 'executing'
                    ? '#10b981, #34d399'
                    : '#f59e0b, #fbbf24'
                })`,
              }}
            />
          </div>
        </div>
      )}

      {/* Chat Area */}
      <div 
        className="flex-1 overflow-y-auto px-3 py-3 bg-gradient-to-b from-gray-50/60 to-white min-h-0 relative"
        style={{ overflowY: 'auto' }}
        ref={(el) => {
          if (el) {
            el.onscroll = () => {
              const diff = el.scrollHeight - el.scrollTop - el.clientHeight;
              setShowScrollDown(diff > 100);
            };
          }
        }}
      >
        {messages.length === 0 ? (
          <WelcomeScreen selectedRange={selectedRange} />
        ) : (
          <>
            {messages.map((msg) => (
              <MessageBubble
                key={msg.id}
                message={msg}
                planSteps={(msg as any).planSteps}
                executionResults={(msg as any).executionResults}
                chartImage={(msg as any).chartImage}
                chartImages={(msg as any).chartImages}
                isReasoning={(msg as any).isReasoning}
              />
            ))}
            {isProcessing && <TypingIndicator phase={processingPhase} />}
          </>
        )}
        <div ref={messagesEndRef} />
        
        {/* Scroll to bottom button */}
        {showScrollDown && (
          <button
            onClick={scrollToBottom}
            className="absolute bottom-4 right-4 w-8 h-8 rounded-full bg-[#217346] text-white shadow-lg flex items-center justify-center hover:bg-[#185C37] active:scale-95 transition-all"
          >
            <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M6 9l6 6 6-6" />
            </svg>
          </button>
        )}
      </div>

      {/* Premium Input */}
      <ChatInput 
        onSend={handleSend} 
        disabled={isProcessing} 
        selectedRange={selectedRange}
        reasoningEnabled={reasoningEnabled}
        onToggleReasoning={handleToggleReasoning}
      />

      {showSettings && (
        <SettingsPanel
          settings={settings}
          onSave={handleSettingsSave}
          onClose={() => setShowSettings(false)}
        />
      )}

      {/* Custom Confirm Modal for Office.js */}
      {ConfirmModalComponent}
    </div>
  );
}

async function executeStep(step: { action: string; params: Record<string, any>; description: string }): Promise<string> {
  const { action, params } = step;
  switch (action) {
    case 'get_workbook_structure': return JSON.stringify(await ExcelAPI.getWorkbookStructure());
    case 'get_selected_range': return JSON.stringify(await ExcelAPI.getSelectedRange());
    case 'get_range': return JSON.stringify(await ExcelAPI.getRange(params.address, params.sheet_name));
    case 'get_sheet_data': return JSON.stringify(await ExcelAPI.getSheetData(params.sheet_name, params.max_rows));
    case 'set_values': 
      console.log('[EXECUTION] set_values called with:', { address: params.address, sheet_name: params.sheet_name, valuesRows: params.values?.length });
      await ExcelAPI.setValues(params.address, params.values, params.sheet_name); 
      return `Values written to ${params.address} in ${params.sheet_name}`;
    case 'create_table': await ExcelAPI.createTable(params.address, params.name, params.sheet_name); return `Created table "${params.name}" at ${params.address}`;
    case 'set_formulas': await ExcelAPI.setFormulas(params.address, params.formulas, params.sheet_name); return `Formulas written to ${params.address}`;
    case 'apply_format': {
      const format = { address: params.address, sheetName: params.sheet_name, bold: params.bold, italic: params.italic, fontColor: params.font_color, fillColor: params.fill_color, fontSize: params.font_size, fontFamily: params.font_family, numberFormat: params.number_format, horizontalAlignment: params.horizontal_alignment, verticalAlignment: params.vertical_alignment, wrapText: params.wrap_text, borders: params.border_all ? { all: true, color: params.border_color || '#000000', style: params.border_style || 'Continuous', weight: params.border_weight || 'Thin' } : undefined };
      console.log('[EXECUTION] apply_format borders:', JSON.stringify(format.borders));
      await ExcelAPI.applyFormat(format);
      return `Formatting applied to ${params.address}`;
    }
    case 'insert_rows': await ExcelAPI.insertRows(params.address, params.count, params.sheet_name); return `Inserted ${params.count} rows at ${params.address}`;
    case 'delete_rows': await ExcelAPI.deleteRows(params.address, params.count, params.sheet_name); return `Deleted ${params.count} rows at ${params.address}`;
    case 'insert_columns': await ExcelAPI.insertColumns(params.address, params.count, params.sheet_name); return `Inserted ${params.count} columns at ${params.address}`;
    case 'delete_columns': await ExcelAPI.deleteColumns(params.address, params.count, params.sheet_name); return `Deleted ${params.count} columns at ${params.address}`;
    case 'add_worksheet': 
      const actualSheetName = await ExcelAPI.addWorksheet(params.name); 
      return `Added worksheet "${actualSheetName}"`;
    case 'delete_worksheet': await ExcelAPI.deleteWorksheet(params.sheet_name || params.name); return `Deleted worksheet "${params.sheet_name || params.name}"`;
    case 'sort_range': await ExcelAPI.sortRange(params.address, params.column_index, params.ascending, params.sheet_name); return `Sorted ${params.address} by column ${params.column_index}`;
    case 'auto_fill': await ExcelAPI.autoFill(params.source_address, params.target_address, params.sheet_name); return `Autofilled from ${params.source_address} to ${params.target_address}`;
    case 'create_chart': return await ExcelAPI.createChart(params.chart_type, params.data_range, params.sheet_name, params.title, params.position);
    case 'conditional_format':
      await ExcelAPI.conditionalFormat(params.address, {
        type: params.rule_type || 'cellValue',
        operator: params.operator,
        formula1: params.formula1,
        formula2: params.formula2,
        format: { fillColor: params.fill_color, fontColor: params.font_color, bold: params.bold },
      }, params.sheet_name);
      return `Conditional formatting applied to ${params.address}`;
    case 'find_replace': {
      const count = await ExcelAPI.findAndReplace(params.address, params.find_text, params.replace_text, params.match_case || false, params.sheet_name);
      return `Replaced ${count} occurrence(s) in ${params.address}`;
    }
    case 'merge_cells': await ExcelAPI.mergeCells(params.address, params.sheet_name); return `Merged cells ${params.address}`;
    case 'unmerge_cells': await ExcelAPI.unmergeCells(params.address, params.sheet_name); return `Unmerged cells ${params.address}`;
    case 'add_dropdown': await ExcelAPI.addDropdown(params.address, params.options, params.sheet_name); return `Added dropdown to ${params.address} with options: ${params.options.join(', ')}`;
    default: throw new Error(`Unknown action: ${action}`);
  }
}

async function getWorkbookContext(): Promise<string | null> {
  try {
    return await new Promise((resolve) => {
      if (typeof Excel === 'undefined') { resolve(null); return; }
      Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load('items/name');
        await context.sync();
        const activeSheet = (context.workbook.worksheets as any).getActiveWorksheet();
        if (activeSheet) { activeSheet.load('name'); await context.sync(); }
        resolve(JSON.stringify({ sheets: sheets.items.map((s: any) => s.name), activeSheet: activeSheet?.name || sheets.items[0]?.name }));
      }).catch(() => resolve(null));
    });
  } catch { return null; }
}

async function getSelectedRangeData(): Promise<string | null> {
  try {
    return await new Promise((resolve) => {
      if (typeof Excel === 'undefined') { resolve(null); return; }
      Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address, rowCount, columnCount, values');
        await context.sync();
        const sheet = range.worksheet;
        sheet.load('name');
        await context.sync();
        
        // Extract headers from first row
        const values = range.values;
        let headers: string[] = [];
        if (values && values.length > 0 && values[0]) {
          for (let i = 0; i < values[0].length; i++) {
            const header = values[0][i];
            headers.push(header ? String(header).trim() : `Column ${i + 1}`);
          }
        }
        
        const result = JSON.stringify({ address: range.address, sheetName: sheet.name, rowCount: range.rowCount, columnCount: range.columnCount, values: range.values, headers: headers });
        console.log('[RANGE] getSelectedRangeData returning:', result.slice(0, 200));
        resolve(result);
      }).catch(() => resolve(null));
    });
  } catch { return null; }
}

function validatePlan(plan: { action: string; params: Record<string, any>; description: string }[], userMessage: string = ''): { validSteps: { action: string; params: Record<string, any>; description: string }[]; validationErrors: string[] } {
  const validSteps: { action: string; params: Record<string, any>; description: string }[] = [];
  const validationErrors: string[] = [];
  const VALID_ACTIONS = new Set(['get_workbook_structure', 'get_selected_range', 'get_range', 'get_sheet_data', 'set_values', 'set_formulas', 'apply_format', 'insert_rows', 'delete_rows', 'insert_columns', 'delete_columns', 'add_worksheet', 'delete_worksheet', 'create_table', 'sort_range', 'auto_fill', 'create_chart', 'conditional_format', 'find_replace', 'merge_cells', 'unmerge_cells', 'add_dropdown']);
  for (let i = 0; i < plan.length; i++) {
    const step = plan[i];
    if (!VALID_ACTIONS.has(step.action)) { validationErrors.push(`Step ${i + 1}: Unknown action "${step.action}" — skipping`); continue; }
    validSteps.push(step);
  }
  return { validSteps, validationErrors };
}
