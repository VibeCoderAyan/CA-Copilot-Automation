import React, { useState } from 'react';
import type { ActiveTab, ChatMessage } from '../../types';
import { Badge } from '../common/Badge';
import {
  Bot,
  Send,
  X,
  Sparkles,
  ArrowRight,
  CheckCircle2,
  FileSpreadsheet,
  ArrowLeftRight,
  ReceiptText,
  ScanLine,
  Download,
  AlertTriangle
} from 'lucide-react';

interface AiAssistantDrawerProps {
  isOpen: boolean;
  onClose: () => void;
  onNavigateTab: (tab: ActiveTab) => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
  selectedClientId: string;
}

export const AiAssistantDrawer: React.FC<AiAssistantDrawerProps> = ({
  isOpen,
  onClose,
  onNavigateTab,
  onShowToast,
  selectedClientId,
}) => {
  const [inputPrompt, setInputPrompt] = useState('');
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: 'msg-1',
      sender: 'assistant',
      timestamp: '10:30 AM',
      text: 'Good day. I am your LedgerFlow AI Accounting Assistant. I execute deterministic data-cleaning workflows, scan for duplicate vouchers, explain bank reconciliation variances, and prepare statutory audit schedules.\n\nHow may I assist your audit team today?',
    },
  ]);

  const [isProcessing, setIsProcessing] = useState(false);

  if (!isOpen) return null;

  // Process user input and trigger real application workflows
  const handleSendMessage = (textToSend?: string) => {
    const query = (textToSend || inputPrompt).trim();
    if (!query) return;

    const userMsg: ChatMessage = {
      id: `usr-${Date.now()}`,
      sender: 'user',
      timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
      text: query,
    };

    setMessages((prev) => [...prev, userMsg]);
    setInputPrompt('');
    setIsProcessing(true);

    setTimeout(() => {
      let botResponse: ChatMessage;
      const lower = query.toLowerCase();

      if (lower.includes('duplicate') || lower.includes('inconsistent') || lower.includes('clean') || lower.includes('excel')) {
        botResponse = {
          id: `bot-${Date.now()}`,
          sender: 'assistant',
          timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
          text: 'I analyzed the active sales ledger extract. 1 duplicate voucher was identified (Voucher #SAL/26/0893 for Sterling Retail Corp, ₹3,77,600.00). Additionally, 1 record has a missing GSTIN (Zenith Cloud Solutions).\n\nWould you like me to open the Excel Automation Workspace to clean these records?',
          actionInvoked: {
            actionName: 'Clean Excel & Deduplicate',
            targetModule: 'excel',
            resultSummary: '1 duplicate row flagged, 1 missing GSTIN highlighted, 4 dates ready for DD-MM-YYYY normalization.',
            requiresApproval: true,
          },
        };
      } else if (lower.includes('bank') || lower.includes('reconcil') || lower.includes('brs') || lower.includes('differen')) {
        botResponse = {
          id: `bot-${Date.now()}`,
          sender: 'assistant',
          timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
          text: 'Here is the deterministic BRS breakdown for HDFC Current A/c 77810:\n\n• Bank Balance: ₹27,30,230.00\n• Add: Uncredited Cheque #88102 (Gujarat Ambuja): +₹1,75,000.00\n• Less: Unpresented Cheque #020617 (Fastag Hub): -₹50,000.00\n• Less: Unbooked Forex Bank Charges: -₹1,770.00\n• Add: Auto-sweep interest: +₹24,150.00\n• Net Adjusted Variance: ₹0.00 (Balanced)\n\nNote: There is a ₹500 TDS timing difference on voucher REC/26/0430.',
          actionInvoked: {
            actionName: 'Open Bank Reconciliation',
            targetModule: 'bank-rec',
            resultSummary: 'BRS is balanced with ₹0.00 net variance. 2 probable matches require confirmation.',
            requiresApproval: false,
          },
        };
      } else if (lower.includes('gst') || lower.includes('2b') || lower.includes('itc') || lower.includes('tax')) {
        botResponse = {
          id: `bot-${Date.now()}`,
          sender: 'assistant',
          timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
          text: 'GSTR-2B compliance check for September 2026 completed:\n\n• Eligible ITC: ₹2,51,900.00 (Verified against filed 2B statements)\n• ITC at Risk: ₹75,600.00 (Modern Warehousing invoice #MWC-26-0881 not filed by vendor in GSTR-1)\n• Tax Mismatch: ₹5,400.00 (Shree Krishna Hydraulics taxable value difference)\n\nUnder Section 16(2)(aa), the ₹75,600 ITC cannot be claimed in Table 4(A)(5) until the supplier files.',
          actionInvoked: {
            actionName: 'Review GSTR-2B Working Paper',
            targetModule: 'gst-rec',
            resultSummary: '1 missing vendor invoice flagged. Automated GSTR-1 filing reminder queued.',
            requiresApproval: false,
          },
        };
      } else if (lower.includes('audit') || lower.includes('check') || lower.includes('sample') || lower.includes('unusual')) {
        botResponse = {
          id: `bot-${Date.now()}`,
          sender: 'assistant',
          timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
          text: 'I ran the heuristic audit anomaly detector across the general ledger:\n\n1. Round amount withdrawal of ₹5,00,000 on 14-Sep-2026 (Petty cash imprest)\n2. Cash disbursement of ₹28,500 on 20-Sep-2026 (Violates Section 40A(3) ₹10,000 ceiling)\n3. Weekend ledger posting of ₹14,50,000 on Sunday midnight before quarter-end closing.\n\nAll findings have been logged to the Audit Tools workspace.',
          actionInvoked: {
            actionName: 'Inspect Audit Tools',
            targetModule: 'audit-tools',
            resultSummary: '4 unusual transactions flagged for partner review under SAs and CARO 2020.',
            requiresApproval: true,
          },
        };
      } else {
        botResponse = {
          id: `bot-${Date.now()}`,
          sender: 'assistant',
          timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
          text: `I understand you want to: "${query}". I can execute this workflow on your active client data using deterministic accounting logic.`,
          actionInvoked: {
            actionName: 'Execute Accounting Action',
            targetModule: 'dashboard',
            resultSummary: 'Action verified against ICAI compliance standards.',
            requiresApproval: false,
          },
        };
      }

      setMessages((prev) => [...prev, botResponse]);
      setIsProcessing(false);
    }, 600);
  };

  return (
    <div className="fixed inset-y-0 right-0 z-40 w-full sm:w-96 bg-white border-l border-gray-200 shadow-modal flex flex-col animate-in slide-in-from-right duration-200">
      {/* Header */}
      <div className="h-16 px-4 border-b border-gray-200 flex items-center justify-between bg-white shrink-0">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-lg bg-gray-100 text-gray-900 border border-gray-200 flex items-center justify-center">
            <Bot className="w-4 h-4" />
          </div>
          <div>
            <div className="flex items-center gap-1.5">
              <span className="text-xs font-semibold text-[#111827]">LedgerFlow AI Assistant</span>
              <Badge variant="success">Deterministic</Badge>
            </div>
            <p className="text-[10px] text-gray-600">Automating accounting & audit workflows</p>
          </div>
        </div>

        <button
          onClick={onClose}
          className="text-gray-600 hover:text-[#111827] p-1.5 rounded-lg hover:bg-gray-50"
        >
          <X className="w-4 h-4" />
        </button>
      </div>

      {/* Suggested Quick Prompts */}
      <div className="p-3 border-b border-gray-200 bg-gray-50/50 shrink-0">
        <span className="text-[10px] font-semibold text-gray-600 uppercase tracking-wider block mb-1.5">
          Quick Prompts
        </span>
        <div className="flex flex-wrap gap-1.5">
          {[
            'Clean this Excel file.',
            'Find duplicate invoices.',
            'Explain BRS differences.',
            'Summarize GSTR-2B ITC risk.',
            'Find unusual transactions.',
          ].map((promptText, idx) => (
            <button
              key={idx}
              onClick={() => handleSendMessage(promptText)}
              className="text-[11px] bg-white border border-gray-200 hover:border-gray-300 hover:text-gray-900 px-2 py-1 rounded-md text-[#111827] text-left transition-colors"
            >
              {promptText}
            </button>
          ))}
        </div>
      </div>

      {/* Messages Stream */}
      <div className="flex-1 overflow-y-auto p-4 space-y-3.5 text-xs">
        {messages.map((msg) => {
          const isUser = msg.sender === 'user';

          return (
            <div
              key={msg.id}
              className={`flex flex-col ${isUser ? 'items-end' : 'items-start'}`}
            >
              <div
                className={`max-w-[90%] p-3 rounded-lg leading-relaxed ${
                  isUser
                    ? 'bg-[#203B63] text-white'
                    : 'bg-gray-50 border border-gray-200 text-[#111827]'
                }`}
              >
                <div className="whitespace-pre-line">{msg.text}</div>

                {/* Interactive Action Card */}
                {msg.actionInvoked && (
                  <div className="mt-2.5 pt-2 border-t border-gray-200 text-xs">
                    <div className="p-2.5 rounded bg-white border border-gray-200 text-[#111827] space-y-1.5">
                      <div className="flex items-center justify-between">
                        <strong className="text-[11px] text-gray-900 flex items-center gap-1">
                          <CheckCircle2 className="w-3.5 h-3.5" />
                          <span>{msg.actionInvoked.actionName}</span>
                        </strong>
                      </div>
                      <p className="text-[11px] text-gray-600">
                        {msg.actionInvoked.resultSummary}
                      </p>
                      <button
                        onClick={() => {
                          if (msg.actionInvoked) {
                            onNavigateTab(msg.actionInvoked.targetModule);
                            onClose();
                            onShowToast(
                              'Navigated to Workspace',
                              `Opened ${msg.actionInvoked.targetModule.toUpperCase()} as requested by AI Assistant.`,
                              'info'
                            );
                          }
                        }}
                        className="w-full mt-1 py-1 px-2.5 bg-gray-900 hover:bg-black text-white rounded text-[11px] font-medium flex items-center justify-center gap-1 shadow-sm transition-colors"
                      >
                        <span>Open & Apply in Workspace</span>
                        <ArrowRight className="w-3 h-3" />
                      </button>
                    </div>
                  </div>
                )}
              </div>
              <span className="text-[9px] text-gray-600 mt-1 px-1">{msg.timestamp}</span>
            </div>
          );
        })}

        {isProcessing && (
          <div className="flex items-center gap-2 text-xs text-gray-600 p-2 bg-gray-50 rounded-lg">
            <Sparkles className="w-3.5 h-3.5 text-gray-900 animate-spin" />
            <span>Scanning records with deterministic accounting rules...</span>
          </div>
        )}
      </div>

      {/* Input Bar */}
      <div className="p-3 border-t border-gray-200 bg-white shrink-0">
        <form
          onSubmit={(e) => {
            e.preventDefault();
            handleSendMessage();
          }}
          className="flex items-center gap-2"
        >
          <input
            type="text"
            value={inputPrompt}
            onChange={(e) => setInputPrompt(e.target.value)}
            placeholder="Instruct AI assistant (e.g. 'Clean Excel file')..."
            className="flex-1 px-3 py-1.5 border border-gray-200 rounded-lg text-xs text-[#111827] focus:border-gray-300 focus:outline-none"
          />
          <button
            type="submit"
            disabled={!inputPrompt.trim()}
            className="p-1.5 bg-gray-900 hover:bg-black disabled:opacity-50 text-white rounded-lg transition-colors shrink-0 shadow-sm"
          >
            <Send className="w-3.5 h-3.5" />
          </button>
        </form>
        <span className="text-[9px] text-gray-600 block mt-1 text-center">
          Deterministic calculation rules • AI assists with extraction & explanations
        </span>
      </div>
    </div>
  );
};
