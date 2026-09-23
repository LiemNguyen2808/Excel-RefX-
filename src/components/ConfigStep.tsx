import React, { useState } from 'react';
import { ColumnMeta, ExtractionConfig, ExtractionPreset } from '../types/excel';
import { CheckSquare, Square, ArrowUpDown, Bookmark, Trash2, ArrowUp, ArrowDown, Search, GitMerge, FileText } from 'lucide-react';

interface ConfigStepProps {
  columns: ColumnMeta[];
  referenceColumnLetter: string;
  onChangeReferenceColumn: (letter: string) => void;
  referenceColumnName?: string;
  selectedColumnLetters: string[];
  onChangeSelectedColumns: (letters: string[]) => void;
  columnAliases: Record<string, string>;
  onChangeColumnAlias: (letter: string, alias: string) => void;
  onReorderColumns: (newOrderedLetters: string[]) => void;
  matchBy: 'letter' | 'header_name';
  onChangeMatchBy: (val: 'letter' | 'header_name') => void;
  includeSourceFileName: boolean;
  onChangeIncludeSourceFileName: (val: boolean) => void;
  includeSheetName: boolean;
  onChangeIncludeSheetName: (val: boolean) => void;
  presets: ExtractionPreset[];
  onSavePreset: (name: string) => void;
  onApplyPreset: (preset: ExtractionPreset) => void;
  onDeletePreset: (presetId: string) => void;
  filesCount: number;
  enabledFilesCount: number;
}

export const ConfigStep: React.FC<ConfigStepProps> = ({
  columns,
  referenceColumnLetter,
  onChangeReferenceColumn,
  referenceColumnName,
  selectedColumnLetters,
  onChangeSelectedColumns,
  columnAliases,
  onChangeColumnAlias,
  onReorderColumns,
  matchBy,
  onChangeMatchBy,
  includeSourceFileName,
  onChangeIncludeSourceFileName,
  includeSheetName,
  onChangeIncludeSheetName,
  presets,
  onSavePreset,
  onApplyPreset,
  onDeletePreset,
  filesCount,
  enabledFilesCount,
}) => {
  const [filterSearch, setFilterSearch] = useState('');
  const [newPresetName, setNewPresetName] = useState('');
  const [showPresetModal, setShowPresetModal] = useState(false);

  // Selected columns set for fast lookup
  const selectedSet = new Set(selectedColumnLetters);

  // Filter columns by user search text
  const filteredColumns = columns.filter(
    (col) =>
      col.letter.toLowerCase().includes(filterSearch.toLowerCase()) ||
      col.name.toLowerCase().includes(filterSearch.toLowerCase()) ||
      String(col.sampleValue || '').toLowerCase().includes(filterSearch.toLowerCase())
  );

  const handleSelectAll = () => {
    onChangeSelectedColumns(columns.map((c) => c.letter));
  };

  const handleDeselectAll = () => {
    onChangeSelectedColumns([]);
  };

  const handleToggleColumn = (letter: string) => {
    if (selectedSet.has(letter)) {
      onChangeSelectedColumns(selectedColumnLetters.filter((l) => l !== letter));
    } else {
      onChangeSelectedColumns([...selectedColumnLetters, letter]);
    }
  };

  const handleMoveColumn = (letter: string, direction: 'up' | 'down') => {
    const idx = selectedColumnLetters.indexOf(letter);
    if (idx === -1) return;
    const targetIdx = direction === 'up' ? idx - 1 : idx + 1;
    if (targetIdx < 0 || targetIdx >= selectedColumnLetters.length) return;

    const updated = [...selectedColumnLetters];
    const [moved] = updated.splice(idx, 1);
    updated.splice(targetIdx, 0, moved);
    onReorderColumns(updated);
  };

  const activeRefCol = columns.find((c) => c.letter === referenceColumnLetter);

  return (
    <div className="bg-white border border-neutral-200 rounded-lg p-5 shadow-xs space-y-5">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-4 border-b border-neutral-100">
        <div>
          <div className="flex items-center gap-2">
            <h2 className="text-base font-semibold text-neutral-900">
              2. Chỉ định Cột Mã Tham Chiếu &amp; Các Cột Xuất Dữ Liệu
            </h2>
            {filesCount > 1 && (
              <span className="text-[11px] font-medium bg-neutral-100 text-neutral-700 px-2 py-0.5 rounded border border-neutral-300">
                Áp dụng cho {enabledFilesCount} tệp đang kích hoạt
              </span>
            )}
          </div>
          <p className="text-xs text-neutral-500 mt-0.5">
            Chọn cột chứa mã cần tra cứu (ví dụ Cột C: PR/ INDENT) và tích chọn các cột muốn lấy ra file mới.
          </p>
        </div>

        {/* Preset Management Button */}
        <div className="flex items-center gap-2">
          {presets.length > 0 && (
            <div className="flex items-center gap-1.5">
              <span className="text-xs text-neutral-500">Mẫu đã lưu:</span>
              <select
                onChange={(e) => {
                  const p = presets.find((pr) => pr.id === e.target.value);
                  if (p) onApplyPreset(p);
                }}
                defaultValue=""
                className="text-xs bg-neutral-50 border border-neutral-300 rounded px-2 py-1 text-neutral-800"
              >
                <option value="" disabled>-- Chọn mẫu --</option>
                {presets.map((p) => (
                  <option key={p.id} value={p.id}>
                    {p.name}
                  </option>
                ))}
              </select>
            </div>
          )}

          <button
            type="button"
            onClick={() => setShowPresetModal(true)}
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-neutral-700 bg-neutral-100 hover:bg-neutral-200 rounded transition-colors"
          >
            <Bookmark className="h-3.5 w-3.5 text-neutral-600" />
            <span>Lưu mẫu cấu hình</span>
          </button>
        </div>
      </div>

      {/* 2A: Reference Column Selection (Primary Key) */}
      <div className="p-4 bg-emerald-50/60 border border-emerald-200 rounded-lg space-y-3">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <label className="text-xs font-bold text-emerald-950 uppercase tracking-wide flex items-center gap-1.5">
              <span className="h-2 w-2 rounded-full bg-emerald-600"></span>
              Cột Mã Tham Chiếu (Reference Key Column)
            </label>
            <p className="text-xs text-emerald-800 mt-0.5">
              Hệ thống sẽ đối chiếu danh sách mã bạn cung cấp với các giá trị tại cột này để lọc ra các dòng dữ liệu.
            </p>
          </div>

          <div className="w-full sm:w-80">
            <select
              value={referenceColumnLetter}
              onChange={(e) => onChangeReferenceColumn(e.target.value)}
              className="w-full text-xs font-semibold text-neutral-900 bg-white border-2 border-emerald-500 rounded-md px-3 py-2 shadow-xs focus:ring-2 focus:ring-emerald-400 focus:outline-none"
            >
              {columns.map((col) => (
                <option key={col.letter} value={col.letter}>
                  Cột {col.letter}: {col.name} {col.sampleValue ? `(Ví dụ: ${String(col.sampleValue).slice(0, 20)})` : ''}
                </option>
              ))}
            </select>
          </div>
        </div>

        {activeRefCol && (
          <div className="pt-2.5 border-t border-emerald-200/80 flex flex-wrap items-center justify-between gap-3 text-xs text-emerald-900">
            <div className="flex flex-wrap items-center gap-4">
              <span>
                Cột đang chọn: <strong className="font-mono bg-white px-1.5 py-0.5 rounded border border-emerald-300 font-bold">{activeRefCol.letter}</strong>
              </span>
              <span>
                Tên tiêu đề: <strong>{activeRefCol.name}</strong>
              </span>
              {activeRefCol.sampleValue && (
                <span>
                  Giá trị mẫu: <strong className="font-mono text-emerald-950 bg-white px-1.5 py-0.5 rounded border border-emerald-200">{String(activeRefCol.sampleValue)}</strong>
                </span>
              )}
            </div>

            {/* Multi-file column resolution mode */}
            {filesCount > 1 && (
              <div className="flex items-center gap-2 bg-white px-2.5 py-1 rounded border border-emerald-300">
                <span className="text-[11px] font-medium text-neutral-700 flex items-center gap-1">
                  <GitMerge className="h-3 w-3 text-emerald-700" />
                  Khi đối chiếu đa tệp:
                </span>
                <select
                  value={matchBy}
                  onChange={(e) => onChangeMatchBy(e.target.value as 'letter' | 'header_name')}
                  className="text-[11px] font-semibold text-neutral-900 bg-transparent border-0 focus:outline-none cursor-pointer"
                >
                  <option value="letter">Cố định theo Cột chữ cái [{activeRefCol.letter}] ở tất cả các tệp</option>
                  <option value="header_name">Tự động tìm cột có tên &quot;{activeRefCol.name}&quot; trong từng tệp</option>
                </select>
              </div>
            )}
          </div>
        )}
      </div>

      {/* Multi-file trace options (Include file name & sheet name) */}
      {filesCount > 1 && (
        <div className="p-3 bg-neutral-50 border border-neutral-200 rounded-md flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center gap-2 text-neutral-700">
            <FileText className="h-4 w-4 text-emerald-600 shrink-0" />
            <span className="font-medium">Tùy chọn cột truy xuất nguồn gốc trong file kết quả mới:</span>
          </div>

          <div className="flex items-center gap-4">
            <label className="flex items-center gap-1.5 cursor-pointer">
              <input
                type="checkbox"
                checked={includeSourceFileName}
                onChange={(e) => onChangeIncludeSourceFileName(e.target.checked)}
                className="rounded border-neutral-300 text-emerald-600 focus:ring-emerald-500"
              />
              <span className="text-neutral-800">Thêm cột &quot;Tệp nguồn&quot;</span>
            </label>

            <label className="flex items-center gap-1.5 cursor-pointer">
              <input
                type="checkbox"
                checked={includeSheetName}
                onChange={(e) => onChangeIncludeSheetName(e.target.checked)}
                className="rounded border-neutral-300 text-emerald-600 focus:ring-emerald-500"
              />
              <span className="text-neutral-800">Thêm cột &quot;Sheet nguồn&quot;</span>
            </label>
          </div>
        </div>
      )}

      {/* 2B: Output Columns Selection */}
      <div className="space-y-3">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
          <div className="flex items-center gap-2">
            <h3 className="text-xs font-bold text-neutral-900 uppercase tracking-wide">
              Các cột cần lấy ra file mới ({selectedColumnLetters.length}/{columns.length} cột đã chọn)
            </h3>
          </div>

          {/* Quick Actions & Search */}
          <div className="flex flex-wrap items-center gap-2">
            <div className="relative">
              <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-neutral-400" />
              <input
                type="text"
                value={filterSearch}
                onChange={(e) => setFilterSearch(e.target.value)}
                placeholder="Lọc tên cột / chữ cái..."
                className="pl-8 pr-2.5 py-1 text-xs bg-neutral-50 border border-neutral-300 rounded focus:bg-white focus:outline-none focus:ring-1 focus:ring-neutral-400 w-48"
              />
            </div>

            <button
              type="button"
              onClick={handleSelectAll}
              className="px-2.5 py-1 text-xs font-medium text-neutral-700 bg-neutral-100 hover:bg-neutral-200 rounded transition-colors"
            >
              Chọn tất cả
            </button>
            <button
              type="button"
              onClick={handleDeselectAll}
              className="px-2.5 py-1 text-xs font-medium text-neutral-700 bg-neutral-100 hover:bg-neutral-200 rounded transition-colors"
            >
              Bỏ chọn tất cả
            </button>
          </div>
        </div>

        {/* Selected Columns Reordering Chips */}
        {selectedColumnLetters.length > 0 && (
          <div className="p-3 bg-neutral-50 border border-neutral-200 rounded-md">
            <div className="flex items-center justify-between pb-2 mb-2 border-b border-neutral-200/80">
              <span className="text-xs font-medium text-neutral-700 flex items-center gap-1.5">
                <ArrowUpDown className="h-3.5 w-3.5 text-neutral-500" />
                Thứ tự các cột sẽ xuất trong file mới (Bấm mũi tên để đổi vị trí cột):
              </span>
              <span className="text-[11px] text-neutral-500">
                Tổng cộng {selectedColumnLetters.length} cột xuất
              </span>
            </div>

            <div className="flex flex-wrap gap-2">
              {selectedColumnLetters.map((letter, idx) => {
                const col = columns.find((c) => c.letter === letter);
                const displayName = columnAliases[letter] || col?.name || `Cột ${letter}`;
                const isRef = letter === referenceColumnLetter;

                return (
                  <div
                    key={letter}
                    className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded text-xs border transition-colors ${
                      isRef
                        ? 'bg-emerald-50 border-emerald-300 text-emerald-950 font-medium'
                        : 'bg-white border-neutral-300 text-neutral-800'
                    }`}
                  >
                    <span className="font-mono font-bold text-neutral-500">{letter}</span>
                    <span className="truncate max-w-36">{displayName}</span>

                    <div className="flex items-center gap-0.5 ml-1 border-l border-neutral-200 pl-1 text-neutral-400">
                      <button
                        type="button"
                        onClick={() => handleMoveColumn(letter, 'up')}
                        disabled={idx === 0}
                        className="hover:text-neutral-900 disabled:opacity-20 p-0.5"
                        title="Di chuyển sang trái"
                      >
                        <ArrowUp className="h-3 w-3 -rotate-90" />
                      </button>
                      <button
                        type="button"
                        onClick={() => handleMoveColumn(letter, 'down')}
                        disabled={idx === selectedColumnLetters.length - 1}
                        className="hover:text-neutral-900 disabled:opacity-20 p-0.5"
                        title="Di chuyển sang phải"
                      >
                        <ArrowDown className="h-3 w-3 -rotate-90" />
                      </button>
                      <button
                        type="button"
                        onClick={() => handleToggleColumn(letter)}
                        className="hover:text-rose-600 p-0.5 ml-0.5 text-neutral-400"
                        title="Xóa cột này khỏi kết quả"
                      >
                        &times;
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* Columns Grid Table */}
        <div className="border border-neutral-200 rounded-md overflow-hidden max-h-72 overflow-y-auto">
          <table className="w-full text-left text-xs border-collapse">
            <thead className="bg-neutral-100 text-neutral-600 sticky top-0 border-b border-neutral-200 z-10">
              <tr>
                <th className="p-2.5 w-12 text-center">Chọn</th>
                <th className="p-2.5 w-16 text-center font-mono">Cột</th>
                <th className="p-2.5 min-w-44">Tên tiêu đề gốc</th>
                <th className="p-2.5 min-w-44">Tên hiển thị khi xuất (Đổi tên nếu muốn)</th>
                <th className="p-2.5 min-w-44">Dữ liệu mẫu</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-neutral-100 bg-white">
              {filteredColumns.map((col) => {
                const isSelected = selectedSet.has(col.letter);
                const isRef = col.letter === referenceColumnLetter;

                return (
                  <tr
                    key={col.letter}
                    className={`hover:bg-neutral-50 transition-colors ${
                      isSelected ? 'bg-neutral-50/60' : ''
                    }`}
                  >
                    <td className="p-2.5 text-center">
                      <button
                        type="button"
                        onClick={() => handleToggleColumn(col.letter)}
                        className="text-neutral-700 hover:text-emerald-600 p-0.5"
                      >
                        {isSelected ? (
                          <CheckSquare className="h-4 w-4 text-emerald-600" />
                        ) : (
                          <Square className="h-4 w-4 text-neutral-400" />
                        )}
                      </button>
                    </td>

                    <td className="p-2.5 text-center font-mono font-bold text-neutral-800">
                      {col.letter}
                      {isRef && (
                        <span className="block text-[9px] text-emerald-700 font-sans font-semibold">
                          [MÃ]
                        </span>
                      )}
                    </td>

                    <td className="p-2.5 text-neutral-900 font-medium">
                      {col.name}
                    </td>

                    <td className="p-2">
                      <input
                        type="text"
                        value={columnAliases[col.letter] ?? ''}
                        placeholder={col.name}
                        onChange={(e) => onChangeColumnAlias(col.letter, e.target.value)}
                        className="w-full text-xs px-2 py-1 border border-neutral-300 rounded focus:outline-none focus:ring-1 focus:ring-neutral-400 bg-white"
                      />
                    </td>

                    <td className="p-2.5 text-neutral-500 font-mono text-[11px] truncate max-w-xs">
                      {col.sampleValue !== null && col.sampleValue !== undefined
                        ? String(col.sampleValue)
                        : <span className="text-neutral-300 italic">(Trống)</span>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Preset Modal */}
      {showPresetModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
          <div className="bg-white rounded-lg p-5 max-w-md w-full shadow-lg border border-neutral-200">
            <h3 className="text-base font-semibold text-neutral-900 mb-1">
              Lưu mẫu cấu hình tra cứu
            </h3>
            <p className="text-xs text-neutral-500 mb-4">
              Lưu lại thiết lập cột tham chiếu và các cột cần xuất để sử dụng lại cho những lần sau chỉ với 1 click.
            </p>

            <div className="space-y-3">
              <div>
                <label className="block text-xs font-medium text-neutral-700 mb-1">
                  Tên mẫu cấu hình:
                </label>
                <input
                  type="text"
                  value={newPresetName}
                  onChange={(e) => setNewPresetName(e.target.value)}
                  placeholder="Ví dụ: Báo cáo Kế hoạch Sợi TCG"
                  className="w-full text-xs px-3 py-2 border border-neutral-300 rounded focus:outline-none focus:ring-1 focus:ring-emerald-500"
                />
              </div>

              {presets.length > 0 && (
                <div className="pt-2 border-t border-neutral-100">
                  <span className="text-xs font-medium text-neutral-700 block mb-1.5">
                    Các mẫu hiện có:
                  </span>
                  <div className="space-y-1 max-h-36 overflow-y-auto">
                    {presets.map((p) => (
                      <div
                        key={p.id}
                        className="flex items-center justify-between p-2 rounded bg-neutral-50 text-xs"
                      >
                        <span className="font-medium text-neutral-800">{p.name}</span>
                        <button
                          type="button"
                          onClick={() => onDeletePreset(p.id)}
                          className="text-neutral-400 hover:text-rose-600 p-1"
                          title="Xóa mẫu này"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            <div className="mt-5 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setShowPresetModal(false)}
                className="px-3 py-1.5 text-xs text-neutral-600 hover:bg-neutral-100 rounded"
              >
                Đóng
              </button>
              <button
                type="button"
                onClick={() => {
                  if (newPresetName.trim()) {
                    onSavePreset(newPresetName.trim());
                    setNewPresetName('');
                    setShowPresetModal(false);
                  }
                }}
                disabled={!newPresetName.trim()}
                className="px-3 py-1.5 text-xs font-medium text-white bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 rounded"
              >
                Lưu mẫu
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
