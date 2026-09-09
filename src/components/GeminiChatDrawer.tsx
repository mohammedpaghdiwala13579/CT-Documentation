import React, { useState, useEffect, useRef } from "react";
import Markdown from "react-markdown";
import { 
  Bot, 
  X, 
  Send, 
  Sparkles, 
  Trash2, 
  Check, 
  Copy, 
  ArrowRight, 
  RefreshCw, 
  FileSearch, 
  Layers, 
  Ship, 
  DollarSign, 
  ShieldCheck, 
  ChevronDown, 
  ListPlus, 
  Minimize2, 
  Maximize2,
  HelpCircle
} from "lucide-react";
import { ChatMessage, GeminiModelChoice, AssistantRole, ExtractedItem, SavedDocument, QuotationRow } from "../types";

interface GeminiChatDrawerProps {
  isOpen: boolean;
  onClose: () => void;
  activeDoc: Partial<SavedDocument> & { rows: QuotationRow[]; grandTotal?: number };
  onImportItemsToQuotation: (items: ExtractedItem[]) => void;
}

const ROLE_DEFINITIONS: Record<AssistantRole, { label: string; icon: any; desc: string; samplePrompts: string[] }> = {
  maritime_analyst: {
    label: "Maritime Operations & RFQ Analyst",
    icon: Ship,
    desc: "IMPA & ISSA codes, deck/engine marine specs, SOLAS safety gear & Bangladeshi port stores.",
    samplePrompts: [
      "Current Chittagong Port fresh provisions prices (beef, chicken, rice, eggs) in BDT & USD",
      "Suggest standard IMPA codes and specs for 24mm mooring ropes & safety harnesses",
      "Price estimate for 220m 28mm PP Mooring Rope & deck stores from Agrabad market",
      "Sourcing genuine marine valves & OEM engine spares from Sitakunda ship breaking yard",
    ],
  },
  pricing_negotiator: {
    label: "Commercial Pricing & Rate Strategist",
    icon: DollarSign,
    desc: "Profit margins, Chittagong Outer Anchorage lighterage launch hire & multi-currency BDT/USD rates.",
    samplePrompts: [
      "Chittagong Outer Anchorage launch boat hire & fresh water barge supply rates",
      "Calculate 22% target markup with anchorage launch delivery costs (in BDT and USD)",
      "Compare USD to BDT tariff buffering for Chittagong port supplies (1 USD ≈ 122 BDT)",
      "Verify Bangladesh NBR zero-rated VAT status for foreign vessel supplies",
    ],
  },
  vessel_auditor: {
    label: "Vessel Compliance & Challan Auditor",
    icon: ShieldCheck,
    desc: "Validates vessel IMO, challan receipts, zero math errors & CPA customs compliance.",
    samplePrompts: [
      "Audit this active quotation for missing vessel details, units, or pricing gaps",
      "Verify delivery challan receipt format for Master sign-off & ship stamp at Chittagong Port",
      "Check export VAT/tax exemption compliance for foreign-flagged vessels",
    ],
  },
};

export default function GeminiChatDrawer({
  isOpen,
  onClose,
  activeDoc,
  onImportItemsToQuotation,
}: GeminiChatDrawerProps) {
  const [messages, setMessages] = useState<ChatMessage[]>(() => {
    try {
      const saved = localStorage.getItem("comilla_gemini_chat_history");
      if (saved) return JSON.parse(saved);
    } catch (e) {
      console.error(e);
    }
    return [
      {
        id: "welcome-1",
        role: "model",
        content: `👋 **Welcome to the Comilla Traders Copilot.**

Powered by Google Gemini with live Bangladeshi market intelligence for **Chittagong (Chattogram) Port, Mongla, Matarbari, and Payra**:
- **Bangladeshi Market & Current Prices**: Real-time pricing for Khatunganj fresh provisions, Agrabad deck/engine stores, Sitakunda ship-breaking OEM spares, and Outer Anchorage launch boat tariffs in **BDT and USD**.
- **Instant RFQ Parsing**: Paste messy captain emails or WhatsApp lists to extract structured quotation rows.
- **IMPA & ISSA Marine Stores**: Suggesting correct marine catalog codes, specifications, and standard units.
- **Document Audit**: Inspecting your active quotation or delivery challan for missing vessel details or mathematical discrepancies.

Ask any question or tap a quick prompt below to get started!`,
        timestamp: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
        modelUsed: "gemini-3.8-flash",
        roleUsed: "maritime_analyst",
      },
    ];
  });

  const [inputPrompt, setInputPrompt] = useState("");
  const [selectedRole, setSelectedRole] = useState<AssistantRole>("maritime_analyst");
  const [selectedModel, setSelectedModel] = useState<GeminiModelChoice>("gemini-3.8-flash");
  const [isLoading, setIsLoading] = useState(false);
  const [isExtracting, setIsExtracting] = useState(false);
  const [extractedPreview, setExtractedPreview] = useState<ExtractedItem[] | null>(null);
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [isExpanded, setIsExpanded] = useState(false);

  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Sync chat history to localStorage
  useEffect(() => {
    try {
      localStorage.setItem("comilla_gemini_chat_history", JSON.stringify(messages));
    } catch (e) {
      console.error(e);
    }
  }, [messages]);

  // Auto-scroll to bottom of messages
  useEffect(() => {
    if (isOpen) {
      messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
    }
  }, [messages, isOpen, isLoading]);

  // Focus textarea when opened
  useEffect(() => {
    if (isOpen) {
      setTimeout(() => textareaRef.current?.focus(), 150);
    }
  }, [isOpen]);

  const handleSendMessage = async (textToSend?: string) => {
    const text = (textToSend || inputPrompt).trim();
    if (!text || isLoading) return;

    const userMessage: ChatMessage = {
      id: `user-${Date.now()}`,
      role: "user",
      content: text,
      timestamp: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
      modelUsed: selectedModel,
      roleUsed: selectedRole,
    };

    const newHistory = [...messages, userMessage];
    setMessages(newHistory);
    if (!textToSend) setInputPrompt("");
    setIsLoading(true);

    try {
      // Build lightweight context of the active document
      const docContext = {
        docType: activeDoc.docType,
        messers: activeDoc.messers,
        vesselName: activeDoc.vesselName,
        address: activeDoc.address,
        requisitionNo: activeDoc.requisitionNo,
        docNumber: activeDoc.challanNo || activeDoc.invoiceNo || "N/A",
        itemCount: activeDoc.rows?.length || 0,
        currency: activeDoc.currency || "USD",
        totalAmount: activeDoc.rows?.reduce((sum, r) => sum + (Number(r.amount) || 0), 0) || 0,
      };

      const res = await fetch("/api/gemini/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          messages: newHistory.map((m) => ({ role: m.role, content: m.content })),
          model: selectedModel,
          role: selectedRole,
          documentContext: docContext,
        }),
      });

      if (!res.ok) {
        const errorData = await res.json().catch(() => ({}));
        throw new Error(errorData.error || `Request failed with status ${res.status}`);
      }

      const data = await res.json();

      const aiMessage: ChatMessage = {
        id: `model-${Date.now()}`,
        role: "model",
        content: data.reply || "No response received from Gemini.",
        timestamp: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
        modelUsed: data.modelUsed || selectedModel,
        roleUsed: data.roleUsed || selectedRole,
      };

      setMessages((prev) => [...prev, aiMessage]);
    } catch (err: any) {
      console.error("Chat error:", err);
      const errorMessage: ChatMessage = {
        id: `err-${Date.now()}`,
        role: "model",
        content: `⚠️ **Unable to complete request:** ${err?.message || "An unexpected error occurred. Please verify your connection or Gemini API key."}`,
        timestamp: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
        modelUsed: selectedModel,
        roleUsed: selectedRole,
      };
      setMessages((prev) => [...prev, errorMessage]);
    } finally {
      setIsLoading(false);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  const handleClearChat = () => {
    if (window.confirm("Clear all conversation history?")) {
      const resetMsg: ChatMessage[] = [
        {
          id: `welcome-${Date.now()}`,
          role: "model",
          content: "Conversation history cleared. Ready for your next maritime supply inquiry.",
          timestamp: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
          modelUsed: selectedModel,
          roleUsed: selectedRole,
        },
      ];
      setMessages(resetMsg);
      localStorage.removeItem("comilla_gemini_chat_history");
    }
  };

  const handleCopyText = (id: string, text: string) => {
    navigator.clipboard.writeText(text);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 1800);
  };

  // Quick Action: Audit Current Document
  const handleAuditDocument = () => {
    const prompt = `Please perform an exhaustive audit of our current active ${activeDoc.docType?.toUpperCase() || "DOCUMENT"}:
- Customer / Messers: ${activeDoc.messers || "Unspecified"}
- Vessel Name: ${activeDoc.vesselName || "Unspecified"}
- Port / Berth: ${activeDoc.portBerth || activeDoc.address || "Unspecified"}
- Total Items: ${activeDoc.rows?.length || 0}
- Requisition No: ${activeDoc.requisitionNo || "None"}
- Challan / Invoice: ${activeDoc.challanNo || activeDoc.invoiceNo || "None"}
Audit the mathematical totals, verify if vital vessel identifiers are missing, and suggest professional enhancements before submitting to the Master.`;
    setSelectedRole("vessel_auditor");
    handleSendMessage(prompt);
  };

  // Extract items from raw text or message content
  const handleExtractItemsFromText = async (text: string) => {
    setIsExtracting(true);
    try {
      const res = await fetch("/api/gemini/extract-items", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text, model: selectedModel }),
      });
      const data = await res.json();
      if (Array.isArray(data.items) && data.items.length > 0) {
        setExtractedPreview(data.items);
      } else {
        alert("No structured maritime items could be identified in the specified text.");
      }
    } catch (e: any) {
      console.error(e);
      alert("Failed to parse items: " + e?.message);
    } finally {
      setIsExtracting(false);
    }
  };

  const handleConfirmImportExtracted = () => {
    if (!extractedPreview || extractedPreview.length === 0) return;
    onImportItemsToQuotation(extractedPreview);
    setExtractedPreview(null);
  };

  if (!isOpen) return null;

  const currentRoleInfo = ROLE_DEFINITIONS[selectedRole];
  const RoleIcon = currentRoleInfo.icon;

  return (
    <div className="fixed inset-0 z-50 flex justify-end bg-slate-900/40 backdrop-blur-xs animate-in fade-in duration-200">
      <div 
        className={`bg-white h-full shadow-2xl flex flex-col border-l border-slate-200 transition-all duration-300 relative ${
          isExpanded ? "w-full md:w-[850px]" : "w-full sm:w-[540px] md:w-[580px]"
        }`}
      >
        {/* Top Header */}
        <header className="px-5 py-3.5 border-b border-slate-200/80 bg-slate-50/90 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-indigo-600 to-indigo-800 text-white flex items-center justify-center shadow-xs">
              <Bot className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-sm font-bold text-slate-900 tracking-tight">Comilla Traders Copilot</h2>
                <span className="bg-emerald-100 text-emerald-800 text-[10px] font-semibold px-2 py-0.5 rounded-full flex items-center gap-1">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                  Online
                </span>
              </div>
              <p className="text-[11px] text-slate-500 font-medium">Enterprise Ship Chandler Assistant</p>
            </div>
          </div>

          <div className="flex items-center gap-1.5">
            <button
              onClick={() => setIsExpanded(!isExpanded)}
              className="text-slate-400 hover:text-slate-700 p-1.5 rounded-lg hover:bg-slate-200/60 transition-colors"
              title={isExpanded ? "Collapse width" : "Expand width"}
            >
              {isExpanded ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
            </button>
            <button
              onClick={handleClearChat}
              className="text-slate-400 hover:text-rose-600 p-1.5 rounded-lg hover:bg-rose-50 transition-colors"
              title="Clear conversation history"
            >
              <Trash2 className="w-4 h-4" />
            </button>
            <button
              onClick={onClose}
              className="text-slate-400 hover:text-slate-700 p-1.5 rounded-lg hover:bg-slate-200/60 transition-colors"
              title="Close drawer"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </header>

        {/* Configuration Bar: Role & Model Selector */}
        <div className="px-5 py-2.5 bg-white border-b border-slate-100 flex flex-wrap items-center justify-between gap-2 shrink-0">
          {/* Role selector */}
          <div className="flex items-center gap-1.5">
            <RoleIcon className="w-3.5 h-3.5 text-indigo-600 shrink-0" />
            <select
              value={selectedRole}
              onChange={(e) => setSelectedRole(e.target.value as AssistantRole)}
              className="text-xs font-semibold text-slate-700 bg-slate-100/80 hover:bg-slate-100 border border-slate-200/80 rounded-lg px-2.5 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer"
            >
              <option value="maritime_analyst">⚓ Operations &amp; RFQ Analyst</option>
              <option value="pricing_negotiator">💰 Pricing &amp; Rate Strategist</option>
              <option value="vessel_auditor">🛡️ Vessel Compliance Auditor</option>
            </select>
          </div>

          {/* Model selector */}
          <div className="flex items-center gap-1 text-[11px]">
            <span className="text-slate-400 font-medium">Model:</span>
            <select
              value={selectedModel}
              onChange={(e) => setSelectedModel(e.target.value as GeminiModelChoice)}
              className="text-[11px] font-semibold text-indigo-900 bg-indigo-50/70 border border-indigo-200/80 rounded-md px-2 py-0.5 focus:outline-none cursor-pointer"
            >
              <option value="gemini-3.8-flash">gemini-3.8-flash (Fast & Accurate)</option>
              <option value="gemini-3.1-flash-lite">gemini-3.1-flash-lite (Ultra-Fast)</option>
              <option value="gemini-3.1-pro-preview">gemini-3.1-pro-preview (Complex Reasoning)</option>
            </select>
          </div>
        </div>

        {/* Active Context Banner */}
        <div className="px-5 py-1.5 bg-slate-50 border-b border-slate-100 flex items-center justify-between text-[11px] text-slate-600">
          <div className="flex items-center gap-2 truncate">
            <span className="font-semibold text-slate-700">Active Doc:</span>
            <span className="truncate bg-white px-2 py-0.5 rounded border border-slate-200 font-mono text-[10px]">
              {activeDoc.vesselName ? `Vessel: ${activeDoc.vesselName}` : activeDoc.messers ? `Client: ${activeDoc.messers}` : "New Quotation"} ({activeDoc.rows?.length || 0} rows)
            </span>
          </div>
          <button
            onClick={handleAuditDocument}
            disabled={isLoading}
            className="text-[11px] text-indigo-600 hover:text-indigo-800 font-semibold flex items-center gap-1 hover:underline cursor-pointer shrink-0"
          >
            <FileSearch className="w-3 h-3" />
            Audit Document
          </button>
        </div>

        {/* Extracted Items Modal / Alert Banner if pending import */}
        {extractedPreview && (
          <div className="p-4 bg-emerald-50 border-b border-emerald-200 text-emerald-900 animate-in slide-in-from-top duration-200">
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-1.5 font-bold text-xs">
                <ListPlus className="w-4 h-4 text-emerald-700" />
                <span>Identified {extractedPreview.length} items from text</span>
              </div>
              <button
                onClick={() => setExtractedPreview(null)}
                className="text-emerald-700 hover:text-emerald-900 text-xs"
              >
                Cancel
              </button>
            </div>
            <div className="max-h-36 overflow-y-auto space-y-1 my-2 pr-1 text-[11px]">
              {extractedPreview.map((item, idx) => (
                <div key={idx} className="flex items-center justify-between bg-white/90 p-1.5 rounded border border-emerald-200/80">
                  <span className="font-medium text-slate-800 truncate max-w-[280px]">{item.desc}</span>
                  <span className="font-mono text-slate-600 text-[10px]">
                    {item.qty || "1"} {item.unit || "PCS"} {item.price ? `@ $${item.price}` : ""}
                  </span>
                </div>
              ))}
            </div>
            <button
              onClick={handleConfirmImportExtracted}
              className="w-full bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs py-1.5 rounded-lg flex items-center justify-center gap-1.5 shadow-sm transition-colors cursor-pointer"
            >
              <Check className="w-3.5 h-3.5" />
              Append {extractedPreview.length} Items to Active Quotation
            </button>
          </div>
        )}

        {/* Scrollable Message Thread */}
        <div className="flex-1 overflow-y-auto p-5 space-y-4">
          {messages.map((message) => {
            const isUser = message.role === "user";
            return (
              <div
                key={message.id}
                className={`flex flex-col ${isUser ? "items-end" : "items-start"} animate-in fade-in duration-150`}
              >
                <div className="flex items-center gap-1.5 mb-1 px-1 text-[10px] text-slate-400">
                  {isUser ? (
                    <span>You</span>
                  ) : (
                    <span className="font-semibold text-indigo-700 flex items-center gap-1">
                      <Bot className="w-3 h-3" />
                      Comilla Traders Copilot
                      {message.modelUsed && (
                        <span className="text-[9px] bg-slate-100 text-slate-500 font-mono px-1 py-0.2 rounded ml-1">
                          {message.modelUsed}
                        </span>
                      )}
                    </span>
                  )}
                  <span>&bull;</span>
                  <span>{message.timestamp}</span>
                </div>

                <div
                  className={`relative group rounded-2xl p-4 max-w-[92%] sm:max-w-[85%] text-xs leading-relaxed shadow-xs ${
                    isUser
                      ? "bg-indigo-600 text-white rounded-tr-xs"
                      : "bg-white text-slate-800 border border-slate-200/90 rounded-tl-xs"
                  }`}
                >
                  {isUser ? (
                    <p className="whitespace-pre-wrap font-sans">{message.content}</p>
                  ) : (
                    <div className="prose prose-xs max-w-none text-slate-800 font-sans space-y-2">
                      <Markdown>{message.content}</Markdown>
                    </div>
                  )}

                  {/* Actions for Assistant Messages */}
                  {!isUser && (
                    <div className="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-[10px] text-slate-400">
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => handleCopyText(message.id, message.content)}
                          className="hover:text-indigo-600 flex items-center gap-1 py-0.5 px-1.5 rounded hover:bg-slate-100 transition-colors"
                        >
                          {copiedId === message.id ? (
                            <>
                              <Check className="w-3 h-3 text-emerald-600" />
                              <span className="text-emerald-600 font-medium">Copied</span>
                            </>
                          ) : (
                            <>
                              <Copy className="w-3 h-3" />
                              <span>Copy</span>
                            </>
                          )}
                        </button>

                        <button
                          onClick={() => handleExtractItemsFromText(message.content)}
                          disabled={isExtracting}
                          className="hover:text-emerald-700 flex items-center gap-1 py-0.5 px-1.5 rounded hover:bg-emerald-50 transition-colors"
                          title="Extract any maritime line items into structured table rows"
                        >
                          <ListPlus className="w-3 h-3 text-emerald-600" />
                          <span>Extract Items to Table</span>
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            );
          })}

          {/* Typing Indicator */}
          {isLoading && (
            <div className="flex items-start gap-2 text-xs text-slate-500">
              <div className="w-7 h-7 rounded-lg bg-indigo-50 border border-indigo-100 text-indigo-600 flex items-center justify-center shrink-0">
                <Bot className="w-4 h-4 animate-spin" />
              </div>
              <div className="bg-white border border-slate-200 rounded-2xl rounded-tl-xs p-3.5 shadow-xs flex items-center gap-2">
                <span className="text-slate-600 font-medium">Gemini is reasoning...</span>
                <span className="flex gap-1">
                  <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-bounce" style={{ animationDelay: "0ms" }} />
                  <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-bounce" style={{ animationDelay: "150ms" }} />
                  <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-bounce" style={{ animationDelay: "300ms" }} />
                </span>
              </div>
            </div>
          )}

          <div ref={messagesEndRef} />
        </div>

        {/* Role Suggestion Pills */}
        <div className="px-5 py-2 bg-slate-50/70 border-t border-slate-100 overflow-x-auto no-scrollbar flex items-center gap-1.5 shrink-0">
          <span className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider shrink-0 flex items-center gap-1">
            <Sparkles className="w-2.5 h-2.5 text-amber-500" />
            Quick:
          </span>
          {currentRoleInfo.samplePrompts.map((prompt, idx) => (
            <button
              key={idx}
              onClick={() => handleSendMessage(prompt)}
              disabled={isLoading}
              className="text-[11px] whitespace-nowrap bg-white hover:bg-indigo-50 text-slate-700 hover:text-indigo-800 border border-slate-200/80 px-2.5 py-1 rounded-full transition-colors cursor-pointer shrink-0 font-medium"
            >
              {prompt.length > 38 ? prompt.substring(0, 38) + "..." : prompt}
            </button>
          ))}
        </div>

        {/* Input Area */}
        <div className="p-4 bg-white border-t border-slate-200/80 shrink-0">
          <div className="relative rounded-xl border border-slate-300 focus-within:border-indigo-500 focus-within:ring-2 focus-within:ring-indigo-100 bg-white transition-all">
            <textarea
              ref={textareaRef}
              rows={2}
              value={inputPrompt}
              onChange={(e) => setInputPrompt(e.target.value)}
              onKeyDown={handleKeyDown}
              placeholder={`Ask the ${currentRoleInfo.label.toLowerCase()} (e.g. Paste RFQ, verify deck stores, estimate margins)...`}
              className="w-full resize-none p-3 pr-12 text-xs text-slate-800 placeholder-slate-400 focus:outline-none bg-transparent"
              disabled={isLoading}
            />

            <button
              onClick={() => handleSendMessage()}
              disabled={!inputPrompt.trim() || isLoading}
              className="absolute right-2.5 bottom-2.5 bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-200 text-white p-1.5 rounded-lg shadow-xs transition-colors cursor-pointer disabled:cursor-not-allowed"
              title="Send message (Enter)"
            >
              <Send className="w-3.5 h-3.5" />
            </button>
          </div>

          <div className="flex items-center justify-between mt-2 px-1 text-[10px] text-slate-400">
            <span>Press <kbd className="font-mono bg-slate-100 px-1 py-0.5 rounded border border-slate-200">Enter</kbd> to send, <kbd className="font-mono bg-slate-100 px-1 py-0.5 rounded border border-slate-200">Shift + Enter</kbd> for newline</span>
            <span>Comilla Traders Copilot</span>
          </div>
        </div>
      </div>
    </div>
  );
}
