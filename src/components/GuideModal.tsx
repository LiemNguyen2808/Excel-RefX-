import React from 'react';
import { X, Layers, FileSpreadsheet, Sparkles, Filter, Download, Zap } from 'lucide-react';

interface GuideModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const GuideModal: React.FC<GuideModalProps> = ({ isOpen, onClose }) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 backdrop-blur-2xs">
      <div className="relative max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-lg bg-white p-6 shadow-xl border border-neutral-200">
        <button
          onClick={onClose}
          className="absolute right-4 top-4 rounded-md p-1 text-neutral-400 hover:bg-neutral-100 hover:text-neutral-700 transition-colors"
        >
          <X className="h-5 w-5" />
        </button>

        <div className="flex items-center gap-2.5 pb-3 border-b border-neutral-200">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-emerald-600 text-white">
            <FileSpreadsheet className="h-5 w-5" />
          </div>
          <div>
            <h3 className="text-base font-bold text-neutral-900">
              Hướng dẫn Sử dụng Excel RefX (Đa tệp &amp; Dung lượng lớn)
            </h3>
            <p className="text-xs text-neutral-500">
              Chuyên xử lý đồng thời nhiều tệp Excel/CSV có cấu trúc phức tạp và kích thước lớn
            </p>
          </div>
        </div>

        <div className="mt-5 space-y-6 text-xs text-neutral-700 leading-relaxed">
          {/* New Feature: Multi-file */}
          <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-md">
            <div className="flex items-center gap-2 text-emerald-950 font-bold text-sm">
              <Zap className="h-4 w-4 text-emerald-600" />
              Nâng cấp: Xử lý Đa tệp cùng lúc &amp; Tối ưu Dung lượng cực lớn
            </div>
            <p className="mt-1 text-emerald-800 text-xs">
              Ứng dụng tích hợp công nghệ <strong>Dense Mode</strong> giúp tiết kiệm đến 80% bộ nhớ RAM, cho phép bạn tải lên cùng lúc hàng chục tệp Excel/CSV với hàng trăm ngàn dòng dữ liệu mà trình duyệt vẫn chạy mượt mà, không bị treo đơ.
            </p>
          </div>

          {/* Step 1 */}
          <div className="flex gap-3">
            <div className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-emerald-100 text-xs font-bold text-emerald-800">
              1
            </div>
            <div>
              <h4 className="font-semibold text-neutral-900 text-sm">
                Nạp hàng loạt tệp &amp; Xử lý cấu trúc phức tạp
              </h4>
              <p className="mt-1 text-neutral-600">
                Bạn có thể kéo thả 1 lúc nhiều tệp Excel (.xlsx, .xls) hoặc CSV (ví dụ: các file theo từng tháng, từng chi nhánh hoặc các nhà cung cấp khác nhau).
              </p>
              <ul className="mt-1.5 list-disc pl-4 space-y-1 text-neutral-600">
                <li>
                  <strong>Bật / Tắt tệp linh hoạt:</strong> Tích chọn các tệp bạn muốn đưa vào lượt tra cứu.
                </li>
                <li>
                  <strong>Dòng tiêu đề &amp; Dòng bắt đầu dữ liệu:</strong> Mỗi tệp có thể có dòng tiêu đề riêng (Dòng 1, Dòng 2...) và dòng dữ liệu riêng (bỏ qua dòng số thứ tự phụ).
                </li>
                <li>
                  <strong>Nút đồng bộ nhanh:</strong> Bấm <em>&quot;Áp dụng dòng này cho TẤT CẢ các tệp&quot;</em> để thiết lập hàng loạt chỉ với 1 click.
                </li>
              </ul>
            </div>
          </div>

          {/* Step 2 */}
          <div className="flex gap-3">
            <div className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-emerald-100 text-xs font-bold text-emerald-800">
              2
            </div>
            <div>
              <h4 className="font-semibold text-neutral-900 text-sm">
                Cơ chế Đối chiếu Đa tệp thông minh
              </h4>
              <p className="mt-1 text-neutral-600">
                Hệ thống hỗ trợ 2 chế độ đối chiếu linh hoạt:
              </p>
              <ul className="mt-1.5 list-disc pl-4 space-y-1 text-neutral-600">
                <li>
                  <strong>Cố định theo chữ cái (Ví dụ Cột C):</strong> Áp dụng cột C ở tất cả các tệp (phù hợp với các biểu mẫu xuất định kỳ từ ERP).
                </li>
                <li>
                  <strong>Tự động tìm theo Tên tiêu đề (Ví dụ &quot;PR/ INDENT&quot;):</strong> Dù tệp A để ở Cột C, tệp B để ở Cột D, hệ thống vẫn tự động tìm chính xác cột có tên tiêu đề tương ứng!
                </li>
              </ul>
            </div>
          </div>

          {/* Step 3 */}
          <div className="flex gap-3">
            <div className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-emerald-100 text-xs font-bold text-emerald-800">
              3
            </div>
            <div>
              <h4 className="font-semibold text-neutral-900 text-sm">
                Chỉ định các cột cần lấy ra &amp; Lưu mẫu (Presets)
              </h4>
              <p className="mt-1 text-neutral-600">
                Tích chọn các cột cần trích xuất (Tên sợi, Nhu cầu, Tồn kho, Khách hàng, PO, Lot...), đổi tên cột và sắp xếp thứ tự tùy ý. Bạn có thể lưu lại mẫu cấu hình để sử dụng cho lần sau.
              </p>
            </div>
          </div>

          {/* Step 4 */}
          <div className="flex gap-3">
            <div className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-emerald-100 text-xs font-bold text-emerald-800">
              4
            </div>
            <div>
              <h4 className="font-semibold text-neutral-900 text-sm">
                Nhập danh sách mã tra cứu &amp; Trích xuất tự động
              </h4>
              <p className="mt-1 text-neutral-600">
                Copy danh sách mã từ Excel và dán vào ô nhập. Hệ thống sẽ quét qua toàn bộ các tệp đã nạp (dù hàng chục vạn dòng dữ liệu chỉ mất chưa tới 1 giây).
              </p>
            </div>
          </div>

          {/* Step 5 */}
          <div className="flex gap-3">
            <div className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-emerald-100 text-xs font-bold text-emerald-800">
              5
            </div>
            <div>
              <h4 className="font-semibold text-neutral-900 text-sm">
                Tổng hợp kết quả ra 1 File duy nhất &amp; Truy xuất nguồn gốc
              </h4>
              <p className="mt-1 text-neutral-600">
                Kết quả từ tất cả các file được tự động gộp chung vào 1 file Excel / CSV mới:
              </p>
              <ul className="mt-1.5 list-disc pl-4 space-y-1 text-neutral-600">
                <li>Tự động gắn cột <strong>&quot;Tệp nguồn&quot;</strong> và <strong>&quot;Sheet nguồn&quot;</strong> để bạn biết chính xác dữ liệu đến từ file nào.</li>
                <li>Tự động tạo sheet <strong>&quot;Thống kê theo tệp&quot;</strong> tổng hợp số dòng khớp ở từng tệp.</li>
                <li>Tự động tạo sheet <strong>&quot;Mã không tìm thấy&quot;</strong> để phục vụ công tác kiểm tra, đối soát.</li>
              </ul>
            </div>
          </div>
        </div>

        <div className="mt-6 flex justify-end pt-4 border-t border-neutral-200">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 text-xs font-semibold text-white bg-neutral-900 hover:bg-neutral-800 rounded-md transition-colors"
          >
            Đã hiểu &amp; Bắt đầu
          </button>
        </div>
      </div>
    </div>
  );
};
