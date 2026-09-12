import React from "react";
import { 
  LayoutDashboard, 
  FileText, 
  FolderKanban, 
  Plus, 
  Ship, 
  ChevronLeft, 
  ChevronRight, 
  Database,
  Truck,
  DollarSign,
  X,
  Building2,
  Wrench,
  Sparkles
} from "lucide-react";
import { CompanyId } from "../types";

export interface ErpSidebarProps {
  activeView: "dashboard" | "editor" | "saved-docs";
  onSelectView: (view: "dashboard" | "editor" | "saved-docs") => void;
  onNewDoc: (type: "quotation" | "challan" | "invoice") => void;
  isCollapsed: boolean;
  onToggleCollapse: () => void;
  isMobileOpen: boolean;
  onCloseMobile: () => void;
  savedDocsCount: number;
  activeCompany?: CompanyId;
  onSelectCompany?: (company: CompanyId) => void;
}

export default function ErpSidebar({
  activeView,
  onSelectView,
  onNewDoc,
  isCollapsed,
  onToggleCollapse,
  isMobileOpen,
  onCloseMobile,
  savedDocsCount,
  activeCompany = "comilla",
  onSelectCompany,
}: ErpSidebarProps) {
  const isZainee = activeCompany === "zainee";

  return (
    <>
      {/* Mobile Backdrop */}
      {isMobileOpen && (
        <div
          id="sidebar-mobile-backdrop"
          onClick={onCloseMobile}
          className="fixed inset-0 bg-slate-950/60 z-40 lg:hidden backdrop-blur-xs transition-opacity"
        />
      )}

      {/* Main Sidebar Container */}
      <aside
        id="erp-sidebar"
        className={`fixed lg:sticky top-0 bottom-0 lg:bottom-auto left-0 lg:left-auto z-40 h-screen bg-[#0B132B] text-slate-300 border-r border-slate-800/80 flex flex-col justify-between transition-all duration-200 shrink-0 no-print print:hidden ${
          isMobileOpen ? "translate-x-0 z-50" : "-translate-x-full lg:translate-x-0"
        } ${isCollapsed ? "w-16" : "w-64"}`}
      >
        {/* Top: Header / Brand Identity */}
        <div>
          <div className="h-14 border-b border-slate-800/80 px-4 flex items-center justify-between">
            {!isCollapsed ? (
              <div className="flex items-center gap-2.5 overflow-hidden">
                <div className="h-8 w-8 rounded-full border border-slate-700 overflow-hidden bg-white flex items-center justify-center shrink-0 shadow-xs">
                  {isZainee ? (
                    <img 
                      src="https://i.ibb.co.com/V8VJdXK/123.png" 
                      alt="ZE Logo" 
                      className="h-full w-full object-contain p-0.5" 
                    />
                  ) : (
                    <div className="h-full w-full bg-blue-600 flex items-center justify-center text-white">
                      <Ship className="h-4 w-4" />
                    </div>
                  )}
                </div>
                <div className="flex flex-col min-w-0">
                  <span className="text-xs font-bold text-white tracking-wider truncate uppercase">
                    {isZainee ? "Zainee Enterprise" : "Comilla Traders"}
                  </span>
                  <span className="text-[10px] text-slate-400 truncate">
                    {isZainee ? "Hardware & Machineries ERP" : "Maritime Supply ERP"}
                  </span>
                </div>
              </div>
            ) : (
              <div className="mx-auto">
                <button
                  type="button"
                  onClick={() => onSelectCompany && onSelectCompany(isZainee ? "comilla" : "zainee")}
                  title={`Switch to ${isZainee ? "Comilla Traders" : "Zainee Enterprise"}`}
                  className="h-8 w-8 rounded-full border border-slate-700 overflow-hidden bg-white flex items-center justify-center shadow-xs cursor-pointer"
                >
                  {isZainee ? (
                    <img 
                      src="https://i.ibb.co.com/V8VJdXK/123.png" 
                      alt="ZE Logo" 
                      className="h-full w-full object-contain p-0.5" 
                    />
                  ) : (
                    <div className="h-full w-full bg-blue-600 flex items-center justify-center text-white">
                      <Ship className="h-4 w-4" />
                    </div>
                  )}
                </button>
              </div>
            )}

            {/* Close button on mobile */}
            <button
              type="button"
              onClick={onCloseMobile}
              className="lg:hidden text-slate-400 hover:text-white p-1 rounded-md cursor-pointer"
              title="Close Navigation"
            >
              <X className="h-4 w-4" />
            </button>
          </div>

          {/* Business Switcher Tabs (Expanded) */}
          {!isCollapsed && onSelectCompany && (
            <div className="px-3 pt-3 pb-2 border-b border-slate-800/60">
              <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider block mb-1.5 px-1">
                Select Business Header
              </span>
              <div className="grid grid-cols-2 gap-1.5 p-1 bg-slate-900/90 rounded-lg border border-slate-800">
                <button
                  type="button"
                  id="btn-sidebar-company-comilla"
                  onClick={() => {
                    onSelectCompany("comilla");
                    onSelectView("editor");
                  }}
                  className={`flex flex-col items-center justify-center py-1.5 px-2 rounded text-[10px] font-semibold transition-all cursor-pointer ${
                    !isZainee
                      ? "bg-blue-600 text-white shadow-xs font-bold"
                      : "text-slate-400 hover:text-white hover:bg-slate-800/70"
                  }`}
                  title="Comilla Traders Marine Supply Header"
                >
                  <span className="truncate">Comilla Traders</span>
                </button>

                <button
                  type="button"
                  id="btn-sidebar-company-zainee"
                  onClick={() => {
                    onSelectCompany("zainee");
                    onSelectView("editor");
                  }}
                  className={`flex flex-col items-center justify-center py-1.5 px-2 rounded text-[10px] font-semibold transition-all cursor-pointer relative ${
                    isZainee
                      ? "bg-emerald-600 text-white shadow-xs font-bold"
                      : "text-slate-400 hover:text-white hover:bg-slate-800/70"
                  }`}
                  title="Zainee Enterprise Hardware & Machineries Header"
                >
                  <span className="truncate flex items-center gap-1">
                    <span>Zainee Ent.</span>
                    <span className="h-1.5 w-1.5 rounded-full bg-emerald-400"></span>
                  </span>
                </button>
              </div>
            </div>
          )}

          {/* Navigation Links */}
          <div className="p-3 space-y-1">
            {/* 1. Dashboard */}
            <button
              type="button"
              id="nav-btn-dashboard"
              onClick={() => {
                onSelectView("dashboard");
                onCloseMobile();
              }}
              className={`w-full flex items-center gap-3 px-3 py-2 rounded-md text-xs font-semibold transition-colors cursor-pointer ${
                activeView === "dashboard"
                  ? "bg-blue-600 text-white shadow-xs"
                  : "text-slate-300 hover:bg-slate-800/60 hover:text-white"
              }`}
              title="Operations Dashboard"
            >
              <LayoutDashboard className="h-4 w-4 shrink-0" />
              {!isCollapsed && <span>Dashboard</span>}
            </button>

            {/* 2. Document Canvas / Editor */}
            <button
              type="button"
              id="nav-btn-editor"
              onClick={() => {
                onSelectView("editor");
                onCloseMobile();
              }}
              className={`w-full flex items-center gap-3 px-3 py-2 rounded-md text-xs font-semibold transition-colors cursor-pointer ${
                activeView === "editor"
                  ? "bg-blue-600 text-white shadow-xs"
                  : "text-slate-300 hover:bg-slate-800/60 hover:text-white"
              }`}
              title="Document Editor (A4 Sheet)"
            >
              <FileText className="h-4 w-4 shrink-0" />
              {!isCollapsed && (
                <div className="flex items-center justify-between w-full min-w-0">
                  <span>Document Canvas</span>
                  <span className="text-[9px] px-1.5 py-0.2 rounded bg-slate-800 text-slate-300 font-mono">
                    {isZainee ? "Zainee" : "Comilla"}
                  </span>
                </div>
              )}
            </button>

            {/* 3. Records Archive */}
            <button
              type="button"
              id="nav-btn-archive"
              onClick={() => {
                onSelectView("saved-docs");
                onCloseMobile();
              }}
              className={`w-full flex items-center justify-between px-3 py-2 rounded-md text-xs font-semibold transition-colors cursor-pointer ${
                activeView === "saved-docs"
                  ? "bg-blue-600 text-white shadow-xs"
                  : "text-slate-300 hover:bg-slate-800/60 hover:text-white"
              }`}
              title="Saved Records & Archive"
            >
              <div className="flex items-center gap-3 min-w-0">
                <FolderKanban className="h-4 w-4 shrink-0" />
                {!isCollapsed && <span className="truncate">Records Archive</span>}
              </div>
              {!isCollapsed && (
                <span className={`text-[10px] font-mono px-1.5 py-0.5 rounded ${
                  activeView === "saved-docs" ? "bg-blue-700 text-white" : "bg-slate-800 text-slate-400"
                }`}>
                  {savedDocsCount}
                </span>
              )}
            </button>
          </div>

          {/* Quick Business Page Jump (Expanded) */}
          {!isCollapsed && onSelectCompany && (
            <div className="px-3 pt-2 mt-1 border-t border-slate-800/60">
              <span className="px-3 text-[10px] font-bold text-slate-500 uppercase tracking-wider block mb-1.5">
                Business Header Pages
              </span>

              <div className="space-y-1">
                <button
                  type="button"
                  onClick={() => {
                    onSelectCompany("comilla");
                    onSelectView("editor");
                    onCloseMobile();
                  }}
                  className={`w-full text-left px-3 py-2 rounded text-[11px] font-medium flex items-center justify-between transition-colors cursor-pointer ${
                    !isZainee
                      ? "bg-blue-900/40 text-blue-200 border border-blue-700/50"
                      : "text-slate-400 hover:text-white hover:bg-slate-800/40"
                  }`}
                >
                  <div className="flex items-center gap-2">
                    <Ship className="h-3.5 w-3.5 text-blue-400 shrink-0" />
                    <span className="truncate">Comilla Traders Page</span>
                  </div>
                  {!isZainee && <span className="text-[9px] bg-blue-600 text-white px-1.5 py-0.2 rounded font-bold">Active</span>}
                </button>

                <button
                  type="button"
                  onClick={() => {
                    onSelectCompany("zainee");
                    onSelectView("editor");
                    onCloseMobile();
                  }}
                  className={`w-full text-left px-3 py-2 rounded text-[11px] font-medium flex items-center justify-between transition-colors cursor-pointer ${
                    isZainee
                      ? "bg-emerald-900/40 text-emerald-200 border border-emerald-700/50"
                      : "text-slate-400 hover:text-white hover:bg-slate-800/40"
                  }`}
                >
                  <div className="flex items-center gap-2">
                    <div className="h-3.5 w-3.5 rounded-full overflow-hidden shrink-0 bg-white">
                      <img src="https://i.ibb.co.com/V8VJdXK/123.png" alt="" className="h-full w-full object-contain" />
                    </div>
                    <span className="truncate">Zainee Enterprise Page</span>
                  </div>
                  {isZainee && <span className="text-[9px] bg-emerald-600 text-white px-1.5 py-0.2 rounded font-bold">Active</span>}
                </button>
              </div>
            </div>
          )}

          {/* Quick Create Section (Expanded only) */}
          {!isCollapsed && (
            <div className="px-3 pt-3 mt-3 border-t border-slate-800/60">
              <span className="px-3 text-[10px] font-bold text-slate-500 uppercase tracking-wider block mb-1.5">
                Quick Create ({isZainee ? "Zainee" : "Comilla"})
              </span>

              <div className="space-y-0.5">
                <button
                  type="button"
                  onClick={() => {
                    onNewDoc("quotation");
                    onCloseMobile();
                  }}
                  className="w-full text-left px-3 py-1.5 rounded text-[11px] font-medium text-slate-400 hover:text-white hover:bg-slate-800/40 flex items-center gap-2 transition-colors cursor-pointer"
                >
                  <Plus className="h-3 w-3 text-blue-400" />
                  <span>Quotation</span>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    onNewDoc("challan");
                    onCloseMobile();
                  }}
                  className="w-full text-left px-3 py-1.5 rounded text-[11px] font-medium text-slate-400 hover:text-white hover:bg-slate-800/40 flex items-center gap-2 transition-colors cursor-pointer"
                >
                  <Truck className="h-3 w-3 text-slate-400" />
                  <span>Challan</span>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    onNewDoc("invoice");
                    onCloseMobile();
                  }}
                  className="w-full text-left px-3 py-1.5 rounded text-[11px] font-medium text-slate-400 hover:text-white hover:bg-slate-800/40 flex items-center gap-2 transition-colors cursor-pointer"
                >
                  <DollarSign className="h-3 w-3 text-emerald-400" />
                  <span>Invoice</span>
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Bottom: Collapse toggle */}
        <div className="p-3 border-t border-slate-800/80">
          {/* Desktop Collapse Trigger */}
          <button
            type="button"
            onClick={onToggleCollapse}
            className="hidden lg:flex w-full items-center justify-center p-1.5 text-slate-400 hover:text-white hover:bg-slate-800/60 rounded-md transition-colors cursor-pointer text-xs"
            title={isCollapsed ? "Expand Sidebar" : "Collapse Sidebar"}
          >
            {isCollapsed ? <ChevronRight className="h-4 w-4" /> : <ChevronLeft className="h-4 w-4" />}
          </button>
        </div>
      </aside>
    </>
  );
}
