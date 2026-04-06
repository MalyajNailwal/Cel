import React, { useState, useCallback, useRef, useEffect } from 'react';
import { cn } from '@/lib/utils';
import { MessageBubble, TypingIndicator, WelcomeScreen } from '@/components/MessageBubble';
import { ChatInput } from '@/components/ChatInput';
import { SettingsPanel } from '@/components/SettingsPanel';
import type { ChatMessage, AIProvider } from '@/lib/ai-providers';
import * as ExcelAPI from '@/lib/excel-api';

type ProcessingPhase = 'idle' | 'thinking' | 'reasoning' | 'planning' | 'analyzing' | 'executing' | 'validating';

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
  thinking: { label: 'Analyzing...', color: 'text-violet-600', accent: 'from-violet-500 to-purple-600' },
  reasoning: { label: 'Thinking...', color: 'text-indigo-600', accent: 'from-indigo-500 to-purple-600' },
  planning: { label: 'Planning...', color: 'text-blue-600', accent: 'from-blue-500 to-indigo-600' },
  analyzing: { label: 'Analyzing data...', color: 'text-cyan-600', accent: 'from-cyan-500 to-blue-600' },
  executing: { label: 'Working...', color: 'text-emerald-600', accent: 'from-emerald-500 to-teal-600' },
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
  const [settings, setSettings] = useState<Settings>(() => {
    try {
      const saved = localStorage.getItem('excel-ai-settings');
      if (saved) return JSON.parse(saved);
    } catch {}
    return { provider: 'openai' as const, model: 'gpt-4o', openaiKey: '', anthropicKey: '', googleKey: '' };
  });

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
      setProcessingPhase('thinking');

      try {
        const workbookContext = await getWorkbookContext();
        const selectedRangeData = await getSelectedRangeData();

        setProcessingPhase('planning');
        const planResponse = await fetch('/api/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            message: userMessage,
            provider: settings.provider,
            model: settings.model,
            api_key: apiKey,
            workbook_context: workbookContext,
            selected_range: selectedRangeData,
            enable_reasoning: reasoningEnabled,
          }),
        });

        if (!planResponse.ok) {
          const text = await planResponse.text();
          throw new Error(`Backend error: ${planResponse.status} - ${text}`);
        }

        const planData = await planResponse.json();
        const plan = planData.plan || [];
        const reasoning = planData.reasoning;
        
        // Show reasoning as a prominent thinking message first
        if (reasoningEnabled && reasoning) {
          setProcessingPhase('reasoning');
          
          // Clean and format reasoning nicely
          const cleanReasoning = reasoning
            .replace(/\*/g, '')
            .replace(/^\d+\.\s*/gm, '')  // Remove numbered list
            .replace(/^-\s*/gm, '')      // Remove bullet points
            .replace(/^•\s*/gm, '')      // Remove bullet symbols
            .replace(/\n{2,}/g, '\n')    // Clean up extra newlines
            .trim();
          
          const reasoningMsg: ExtendedMessage = {
            id: `reasoning-${Date.now()}`,
            role: 'assistant',
            content: cleanReasoning,
            isReasoning: true,
          };
          setMessages((prev) => [...prev, reasoningMsg]);
          setTimeout(scrollToBottom, 100);
        }
        
        const { validSteps, validationErrors } = validatePlan(plan);

        // Parse selected range data for context awareness
        let selectedRange: { address: string; sheetName: string; rowCount: number; columnCount: number } | null = null;
        if (selectedRangeData) {
          try {
            const sr = JSON.parse(selectedRangeData);
            if (sr.address && sr.sheetName) {
              selectedRange = { address: sr.address, sheetName: sr.sheetName, rowCount: sr.rowCount || 1, columnCount: sr.columnCount || 1 };
            }
          } catch {}
        }

        // Check if user is referring to selected area
        const userWantsSelected = /selected|these|this|here|highlighted|current|this range|these cells/i.test(userMessage);

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
          setProcessingPhase('analyzing');
          let selectedData = null;
          try {
            const sr = JSON.parse(selectedRangeData);
            if (sr.values && sr.values.length > 0) {
              selectedData = sr.values;
            }
          } catch {}

          if (selectedData) {
            const isLargeData = selectedData.length > 1000;
            const endpoint = isLargeData ? '/api/analyze-large' : '/api/analyze';
            
            const analyzeResponse = await fetch(endpoint, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                data: selectedData,
                question: userMessage,
                provider: settings.provider,
                model: settings.model,
                api_key: apiKey,
                headers: selectedData[0] || [],
              }),
            });

            if (analyzeResponse.ok) {
              const analysisResult = await analyzeResponse.json();

              if (isChartRequest && selectedRangeData) {
                try {
                  const sr = JSON.parse(selectedRangeData);
                  const headers = selectedData[0] || [];
                  const rows = selectedData.slice(1);
                  const address = sr.address;
                  const sheetName = sr.sheetName;
                  let chartsCreated = 0;
                  let chartDesc = '';

                  const addrParts = address.split('!');
                  const rangePart = addrParts.length > 1 ? addrParts[1] : addrParts[0];
                  const rangeCells = rangePart.split(':');
                  const startCell = rangeCells[0];
                  const endCell = rangeCells.length > 1 ? rangeCells[1] : rangeCells[0];
                  const startColMatch = startCell.match(/^([A-Z]+)/);
                  const endColMatch = endCell.match(/^([A-Z]+)/);
                  const startColIdx = startColMatch ? startColMatch[1].charCodeAt(0) - 65 : 0;
                  const endColIdx = endColMatch ? endColMatch[1].charCodeAt(0) - 65 : startColIdx;
                  const numCols = headers.length;

                  const numericCols: number[] = [];
                  const categoricalCols: number[] = [];
                  for (let i = 0; i < headers.length; i++) {
                    const h = headers[i].toLowerCase();
                    if (/name|id|blood|gender|status|department|region|product|city|category|activity|date|month/i.test(h)) {
                      categoricalCols.push(i);
                    } else if (/age|hemoglobin|rbc|wbc|platelets|salary|revenue|score|quantity|price|count|amount|total|bp|sugar|cholesterol|triglycerides|creatinine|urea|sgot|sgpt|tsh|vitamin|iron|ferritin|hours|sales|units/i.test(h)) {
                      numericCols.push(i);
                    } else {
                      const sampleVal = rows[Math.min(2, rows.length - 1)]?.[i];
                      if (sampleVal !== undefined && sampleVal !== null) {
                        const isNum = !isNaN(Number(String(sampleVal).replace(/[^0-9.-]/g, '')));
                        if (isNum) numericCols.push(i);
                        else categoricalCols.push(i);
                      }
                    }
                  }

                  const startRow = startCell.match(/\d+/)?.[0] || '1';
                  const endRow = endCell.match(/(\d+)$/)?.[1] || '100';

                  const createChartForColumn = async (colIdx: number, chartType: string, title: string) => {
                    const colLetter = String.fromCharCode(65 + startColIdx + colIdx);
                    const chartRange = `${colLetter}${startRow}:${colLetter}${endRow}`;
                    await ExcelAPI.createChart(chartType, chartRange, sheetName, title);
                    chartsCreated++;
                    chartDesc += `📊 ${chartType} chart: ${title}\n`;
                  };

                  const createChartWithTwoColumns = async (col1Idx: number, col2Idx: number, chartType: string, title: string) => {
                    const minCol = startColIdx + Math.min(col1Idx, col2Idx);
                    const maxCol = startColIdx + Math.max(col1Idx, col2Idx);
                    const startLetter = String.fromCharCode(65 + minCol);
                    const endLetter = String.fromCharCode(65 + maxCol);
                    const chartRange = `${startLetter}${startRow}:${endLetter}${endRow}`;
                    await ExcelAPI.createChart(chartType, chartRange, sheetName, title);
                    chartsCreated++;
                    chartDesc += `📊 ${chartType} chart: ${title}\n`;
                  };

                  if (isPieRequest) {
                    if (categoricalCols.length > 0 && numericCols.length > 0) {
                      await createChartWithTwoColumns(categoricalCols[0], numericCols[0], 'pie', `${headers[categoricalCols[0]]} Distribution`);
                    } else if (categoricalCols.length > 0 && numCols >= 2) {
                      await createChartWithTwoColumns(categoricalCols[0], 0, 'pie', `${headers[categoricalCols[0]]} Count`);
                    } else if (numCols >= 2) {
                      await createChartWithTwoColumns(0, 1, 'pie', `${headers[0]} Distribution`);
                    } else {
                      await createChartForColumn(0, 'pie', `${headers[0]} Distribution`);
                    }
                  } else if (isBarRequest) {
                    if (numericCols.length > 0 && categoricalCols.length > 0) {
                      await createChartWithTwoColumns(categoricalCols[0], numericCols[0], 'column', `${headers[numericCols[0]]} by ${headers[categoricalCols[0]]}`);
                    } else if (numericCols.length > 0) {
                      await createChartForColumn(numericCols[0], 'column', `${headers[numericCols[0]]} Distribution`);
                    } else if (numCols >= 2) {
                      await createChartWithTwoColumns(0, 1, 'column', `${headers[1]} Distribution`);
                    } else {
                      await createChartForColumn(0, 'column', `${headers[0]} Distribution`);
                    }
                  } else if (isLineRequest) {
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
                  } else {
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

                  const aiMsg: ExtendedMessage = {
                    id: `msg-${Date.now()}`,
                    role: 'assistant',
                    content: `${analysisResult.analysis || 'Analysis complete.'}\n\n${chartDesc}`,
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

        setProcessingPhase('executing');
        const executionResults: { action: string; success: boolean; output: string }[] = [];
        let lastCreatedSheet: string | null = null;
        let lastDeletedSheet: string | null = null;
        const knownSheets = new Set<string>();

        if (workbookContext) {
          try {
            const ctx = JSON.parse(workbookContext);
            ctx.sheets?.forEach((s: string) => knownSheets.add(s));
          } catch {}
        }

        for (let i = 0; i < validSteps.length; i++) {
          const step = validSteps[i];
          const params = { ...step.params };

          // CRITICAL: Force selected range for write operations when user refers to selection
          if (userWantsSelected && selectedRange) {
            const writeActions = ['set_values', 'set_formulas', 'apply_format', 'create_chart', 'create_table', 'sort_range', 'auto_fill'];
            if (writeActions.includes(step.action)) {
              if (!params.address || params.address !== selectedRange.address) {
                params.address = selectedRange.address;
                validationErrors.push(`Step ${i + 1}: Using selected range "${selectedRange.address}" as target`);
              }
              if (!params.sheet_name || params.sheet_name !== selectedRange.sheetName) {
                params.sheet_name = selectedRange.sheetName;
                validationErrors.push(`Step ${i + 1}: Using selected sheet "${selectedRange.sheetName}"`);
              }
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
            
            const rows = params.values.length;
            const cols = params.values[0]?.length || 1;
            const colLetter = cols <= 26 ? String.fromCharCode(64 + cols) : String.fromCharCode(64 + Math.floor((cols - 1) / 26)) + String.fromCharCode(64 + ((cols - 1) % 26) + 1);
            const correctAddress = `A1:${colLetter}${rows}`;
            
            if (params.address && params.address !== correctAddress) {
              validationErrors.push(`Step ${i + 1}: Address "${params.address}" doesn't match data size — corrected to "${correctAddress}"`);
            }
            params.address = correctAddress;
            
            if (!params.sheet_name && lastCreatedSheet) {
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
          }

          if (step.action === 'apply_format') {
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
            if (!params.address) {
              params.address = 'A1';
              validationErrors.push(`Step ${i + 1}: Missing address — using "A1"`);
            }
            if (!params.sheet_name && lastCreatedSheet) {
              params.sheet_name = lastCreatedSheet;
              validationErrors.push(`Step ${i + 1}: Using sheet "${lastCreatedSheet}" for table`);
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
              if (step.action === 'add_worksheet') {
                const actualName = successOutput.match(/"([^"]+)"/)?.[1] || params.name;
                lastCreatedSheet = actualName;
                knownSheets.add(actualName);
                validationErrors.push(`Step ${i + 1}: Updated lastCreatedSheet to "${actualName}"`);
              }
            } else {
              executionResults.push({ action: step.action, success: false, output: lastError });
            }
          } catch (error) {
            const errMsg = error instanceof Error ? error.message : String(error);
            executionResults.push({ action: step.action, success: false, output: errMsg });
          }
        }

        setProcessingPhase('validating');
        const validationResponse = await fetch('/api/execute', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            message: userMessage,
            provider: settings.provider,
            model: settings.model,
            api_key: apiKey,
            results: executionResults,
          }),
        });

        let finalContent = planData.content;
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
          <span className={cn('text-[11px] font-semibold', phase.color)}>{phase.label}</span>
          <div className="flex-1 h-1 bg-gray-200/60 rounded-full overflow-hidden">
            <div
              className={cn('h-full rounded-full transition-all duration-500')}
              style={{
                width: processingPhase === 'thinking' || processingPhase === 'reasoning' ? '25%' : processingPhase === 'planning' ? '50%' : processingPhase === 'executing' ? '75%' : '95%',
                background: `linear-gradient(90deg, ${
                  processingPhase === 'thinking' || processingPhase === 'reasoning' ? '#8b5cf6, #a78bfa' :
                  processingPhase === 'planning' ? '#3b82f6, #60a5fa' :
                  processingPhase === 'executing' ? '#10b981, #34d399' :
                  '#f59e0b, #fbbf24'
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
          <WelcomeScreen />
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
    case 'set_values': await ExcelAPI.setValues(params.address, params.values, params.sheet_name); return `Values written to ${params.address}`;
    case 'set_formulas': await ExcelAPI.setFormulas(params.address, params.formulas, params.sheet_name); return `Formulas written to ${params.address}`;
    case 'apply_format': {
      const format = { address: params.address, sheetName: params.sheet_name, bold: params.bold, italic: params.italic, fontColor: params.font_color, fillColor: params.fill_color, fontSize: params.font_size, fontFamily: params.font_family, numberFormat: params.number_format, horizontalAlignment: params.horizontal_alignment, verticalAlignment: params.vertical_alignment, wrapText: params.wrap_text };
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
    case 'create_table': await ExcelAPI.createTable(params.address, params.name, params.sheet_name); return `Created table "${params.name}" at ${params.address}`;
    case 'sort_range': await ExcelAPI.sortRange(params.address, params.column_index, params.ascending, params.sheet_name); return `Sorted ${params.address} by column ${params.column_index}`;
    case 'auto_fill': await ExcelAPI.autoFill(params.source_address, params.target_address, params.sheet_name); return `Autofilled from ${params.source_address} to ${params.target_address}`;
    case 'create_chart': return await ExcelAPI.createChart(params.chart_type, params.data_range, params.sheet_name, params.title, params.position);
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
        resolve(JSON.stringify({ address: range.address, sheetName: sheet.name, rowCount: range.rowCount, columnCount: range.columnCount, values: range.values }));
      }).catch(() => resolve(null));
    });
  } catch { return null; }
}

function validatePlan(plan: { action: string; params: Record<string, any>; description: string }[]): { validSteps: { action: string; params: Record<string, any>; description: string }[]; validationErrors: string[] } {
  const validSteps: { action: string; params: Record<string, any>; description: string }[] = [];
  const validationErrors: string[] = [];
  const VALID_ACTIONS = new Set(['get_workbook_structure', 'get_selected_range', 'get_range', 'get_sheet_data', 'set_values', 'set_formulas', 'apply_format', 'insert_rows', 'delete_rows', 'insert_columns', 'delete_columns', 'add_worksheet', 'delete_worksheet', 'create_table', 'sort_range', 'auto_fill', 'create_chart']);
  for (let i = 0; i < plan.length; i++) {
    const step = plan[i];
    if (!VALID_ACTIONS.has(step.action)) { validationErrors.push(`Step ${i + 1}: Unknown action "${step.action}" — skipping`); continue; }
    validSteps.push(step);
  }
  return { validSteps, validationErrors };
}
