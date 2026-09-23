import * as XLSX from 'xlsx';

export const SAMPLE_QUERY_CODES = [
  'TCG-120895-001',
  'TCG-120899-001',
  'TCG-120904-001',
  'TCG-120906-001',
  'TCG-120915-002',
  'TCG-120920-003',
  'TCG-120933-005',
  'TCG-120940-008',
  'TCG-120999-NOTFOUND', // deliberately unmatched to show report
];

const headers = [
  'NOTE',
  'PR Date.YEAR',
  'PR/ INDENT',
  'YARN KIND',
  'code',
  'YARN DEMAND (KG)',
  'STATUS IN W/H (KG)',
  'ON THE WAY TO W/H',
  'USED (KG)',
  'REMAINED/STOCK (KG)',
  'BUYER',
  'SALES',
  'P/O',
  'SCHEDULE YARN',
  'SUPLIER',
  'LOT',
  'LOẠI HÌNH',
  'TESTING RESULT (YES/NO)',
  'GR',
];

const subHeader = [
  1, 2, 3, '', '', '', '7=', 8, 9, 10, '', '', 11, 12, 13, 14, 15, 16, '',
];

export function createSampleWorkbook(monthNumber: number = 9): ArrayBuffer {
  let sampleRows: (string | number | null)[][] = [];

  if (monthNumber === 8) {
    sampleRows = [
      headers,
      subHeader,
      [
        'Tháng 8 đợt 1', '10-08-26', 'TCG-120895-001', 'CVC 30/1 (50/50) (RENU 50%) MVS', '',
        23.91, 24.3, -0.39, 24.3, '-', 'THANHCONG TEX', 'thanh cong',
        'RBD10-TN-LO-0826-8012.1', '25/08/2026', 'SOI 1', 'W-H26-08-A12', 'ID', 'YES', 'GR-08A',
      ],
      [
        '', '12-08-26', 'TCG-120899-001', 'CVC 30/1 (50/50) (RENU 50%) MVS', '',
        23.91, 24.3, -0.39, 24.3, '-', 'THANHCONG TEX', 'thanh cong',
        'RBD10-TN-LO-0826-8012.2', '28/08/2026', 'SOI 1', 'W-H26-08-A12', 'ID', 'YES', 'GR-08A',
      ],
      [
        'Bổ sung', '15-08-26', 'TCG-120895-001', 'CVC 30/1 (50/50) (RENU 50%) MVS', '',
        15.00, 15.0, 0, 15.0, 0, 'THANHCONG TEX', 'thanh cong',
        'RBD10-TN-LO-0826-8012.3', '30/08/2026', 'SOI 1', 'W-H26-08-A12', 'ID', 'YES', 'GR-08A',
      ],
    ];
  } else if (monthNumber === 10) {
    sampleRows = [
      headers,
      subHeader,
      [
        '', '05-10-26', 'TCG-120933-005', 'POLY 100% 75D/72F DTY', 'PLY75',
        88.40, 90.0, -1.6, 90.0, 0, 'HANAE CO', 'thanh cong',
        'HN-TX-8802', '20/10/2026', 'FORMOSA', 'FM-LOT-910', 'ID', 'YES', 'GR-05',
      ],
      [
        'Giao kho số 2', '08-10-26', 'TCG-120940-008', 'CVC 45/55 32/1 SIRO', 'CVC32',
        45.80, 46.0, -0.2, 40.0, 6.0, 'THANHCONG TEX', 'minh quan',
        'RBD10-TN-LO-1026-9070', '22/10/2026', 'SOI 1', 'W-H26-10-B10', 'ID', 'YES', 'GR-01',
      ],
      [
        '', '12-10-26', 'TCG-120940-008', 'CVC 45/55 32/1 SIRO', 'CVC32',
        50.00, 50.0, 0, 50.0, 0, 'THANHCONG TEX', 'minh quan',
        'RBD10-TN-LO-1026-9071', '25/10/2026', 'SOI 1', 'W-H26-10-B10', 'ID', 'YES', 'GR-01',
      ],
    ];
  } else {
    // Default Month 9
    sampleRows = [
      headers,
      subHeader,
      [
        '', '16-09-26', 'TCG-120895-001', 'CVC 30/1 (50/50) (RENU 50%) MVS', '',
        23.91, 24.3, -0.39, 24.3, '-', 'THANHCONG TEX', 'thanh cong',
        'RBD10-TN-LO-0926-9058.1', '30/09/2026', 'SOI 1', 'W-H26-09-A32', 'ID', 'YES', 'GR-01',
      ],
      [
        '', '16-09-26', 'TCG-120899-001', 'CVC 30/1 (50/50) (RENU 50%) MVS', '',
        23.91, 24.3, -0.39, 24.3, '-', 'THANHCONG TEX', 'thanh cong',
        'RBD10-TN-LO-0926-9058.2', '30/09/2026', 'SOI 1', 'W-H26-09-A32', 'ID', 'YES', 'GR-01',
      ],
      [
        '', '16-09-26', 'TCG-120904-001', 'CVC 30/1 (50/50) (RENU 50%) MVS', '',
        19.21, 19.5, -0.29, 19.5, '-', 'THANHCONG TEX', 'thanh cong',
        'RBD10-TN-LO-0926-9059.1', '30/09/2026', 'SOI 1', 'W-H26-09-A32', 'ID', 'YES', 'GR-02',
      ],
      [
        '', '16-09-26', 'TCG-120906-001', 'CVC 30/1 (50/50) (RENU 50%) MVS', '',
        19.21, 19.5, -0.29, 19.5, '-', 'THANHCONG TEX', 'thanh cong',
        'RBD10-TN-LO-0926-9059.2', '30/09/2026', 'SOI 1', 'W-H26-09-A32', 'ID', 'YES', 'GR-02',
      ],
      [
        'Ghi chú bổ sung', '17-09-26', 'TCG-120915-002', 'COTTON 40/1 COMBED COMPACT', 'CT40',
        54.10, 55.0, -0.90, 55.0, 0, 'GLOBAL TEXTILE', 'ngoc mai',
        'PO-GLOBAL-2026-01', '05/10/2026', 'SOI PHUBAI', 'PB-LOT-8821', 'FOB', 'YES', 'GR-03',
      ],
      [
        'Giao đợt 2', '17-09-26', 'TCG-120915-002', 'COTTON 40/1 COMBED COMPACT', 'CT40',
        32.50, 32.5, 0, 32.5, 0, 'GLOBAL TEXTILE', 'ngoc mai',
        'PO-GLOBAL-2026-02', '12/10/2026', 'SOI PHUBAI', 'PB-LOT-8822', 'FOB', 'YES', 'GR-03',
      ],
      [
        '', '18-09-26', 'TCG-120920-003', 'TC 65/35 20/1 CARDED', 'TC20',
        120.00, 125.0, -5.0, 100.0, 25.0, 'SAE-A INTL', 'anh tuan',
        'SAEA-VN-4491', '15/10/2026', 'NAM DINH YARN', 'ND-2026-L4', 'CMT', 'NO', 'GR-04',
      ],
    ];
  }

  const ws = XLSX.utils.aoa_to_sheet(sampleRows);
  const wb = XLSX.utils.book_new();
  const sheetTitle = `Ke_Hoach_Soi_T${monthNumber < 10 ? '0' + monthNumber : monthNumber}`;
  XLSX.utils.book_append_sheet(wb, ws, sheetTitle);

  return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
}

export function createMultiSampleWorkbooks(): { name: string; buffer: ArrayBuffer; size: number }[] {
  return [
    {
      name: 'Ke_Hoach_Soi_Thang_08_2026.xlsx',
      buffer: createSampleWorkbook(8),
      size: 14200,
    },
    {
      name: 'Ke_Hoach_Soi_Thang_09_2026.xlsx',
      buffer: createSampleWorkbook(9),
      size: 18500,
    },
    {
      name: 'Ke_Hoach_Soi_Thang_10_2026.xlsx',
      buffer: createSampleWorkbook(10),
      size: 15300,
    },
  ];
}
