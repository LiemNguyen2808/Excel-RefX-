import React from 'react';
import { FileSpreadsheet, Sparkles, BookOpen, RotateCcw } from 'lucide-react';

interface TopBarProps {
  onLoadSample: () => void;
  onOpenGuide: () => void;
  onReset: () => void;
  hasData: boolean;
}

export const TopBar: React.FC<TopBarProps> = ({
  onLoadSample,
  onOpenGuide,
  onReset,
  hasData,
}) => {
  return (
    <header className="sticky top-0 z-40 w-full border-b border-neutral-200 bg-white/95 backdrop-blur-xs">
      <div className="mx-auto flex h-14 max-w-7xl items-center justify-between px-4 sm:px-6 lg:px-8">
        {/* Zone 1: Single text element wordmark */}
        <div className="flex items-center gap-3">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-emerald-600 text-white shadow-xs">
            <FileSpreadsheet className="h-5 w-5" />
          </div>
          <div>
            <span className="text-base font-bold tracking-tight text-neutral-900">
              Excel RefX
            </span>
            <span className="hidden text-xs text-neutral-500 sm:inline-block sm:ml-2">
              Trích xuất & Tham chiếu cột tự động
            </span>
          </div>
        </div>

        {/* Zone 2: Navigation & Status indicator */}
        <div className="hidden md:flex items-center gap-2 text-xs font-medium text-neutral-500">
          <span className="flex items-center gap-1.5 text-neutral-800">
            <span className="flex h-5 w-5 items-center justify-center rounded-full bg-neutral-900 text-[10px] text-white">1</span>
            Nạp tệp nguồn
          </span>
          <span className="text-neutral-300">/</span>
          <span className={`flex items-center gap-1.5 ${hasData ? 'text-neutral-800' : 'text-neutral-400'}`}>
            <span className={`flex h-5 w-5 items-center justify-center rounded-full text-[10px] ${hasData ? 'bg-neutral-900 text-white' : 'bg-neutral-200 text-neutral-600'}`}>2</span>
            Cột & Mã
          </span>
          <span className="text-neutral-300">/</span>
          <span className={`flex items-center gap-1.5 ${hasData ? 'text-neutral-800' : 'text-neutral-400'}`}>
            <span className={`flex h-5 w-5 items-center justify-center rounded-full text-[10px] ${hasData ? 'bg-neutral-900 text-white' : 'bg-neutral-200 text-neutral-600'}`}>3</span>
            Trích xuất & Xuất file
          </span>
        </div>

        {/* Zone 3: Primary actions */}
        <div className="flex items-center gap-2">
          {hasData && (
            <button
              onClick={onReset}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-neutral-700 bg-neutral-100 hover:bg-neutral-200 rounded-md transition-colors whitespace-nowrap"
              title="Tải tệp mới"
            >
              <RotateCcw className="h-3.5 w-3.5" />
              <span>Nạp lại</span>
            </button>
          )}

          <button
            onClick={onLoadSample}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded-md transition-colors whitespace-nowrap"
            title="Nạp dữ liệu dệt may mẫu theo hình minh họa"
          >
            <Sparkles className="h-3.5 w-3.5 text-emerald-600" />
            <span>Thử dữ liệu mẫu (Mã sợi CVC)</span>
          </button>

          <button
            onClick={onOpenGuide}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-neutral-700 hover:text-neutral-900 border border-neutral-300 hover:border-neutral-400 rounded-md transition-colors whitespace-nowrap"
          >
            <BookOpen className="h-3.5 w-3.5 text-neutral-500" />
            <span>Hướng dẫn</span>
          </button>
        </div>
      </div>
    </header>
  );
};
