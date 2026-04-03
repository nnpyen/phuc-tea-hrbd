# HR + Business Development Dashboard 2026 (Web Prototype)

## Run

1. Open directly:
   - `index.html`
2. Or run local server:
   - `cd /Users/nguyenngocphiyen/Documents/Playground/hrbd-dashboard-2026-web`
   - `python3 -m http.server 8080`
   - Open `http://localhost:8080`

## Data source

- Workbook source: `/Users/nguyenngocphiyen/Downloads/KẾ HOẠCH 2026.xlsx`
- Extract script:
  - `python3 scripts/extract_plan_2026.py`
- Output:
  - `data/plan_2026.json`
  - `data/plan_2026_data.js`

## Scope implemented

- Sidebar navigation for all 8 requested pages.
- Dashboard overview with:
  - Annual OKR summary
  - Target vs Actual cumulative (month / quarter / YTD)
  - Target-weighted achievement rate
  - Year-end forecast
  - Insight blocks: nhận định, nhận xét, phân tích, đề xuất hành động
- Quarter plan page with KPI table + quarter analysis.
- Pillar pages:
  - HR
  - Franchise VN
  - Franchise International
- Import center page:
  - 6 khối nhập liệu:
    - OKR năm theo pillar
    - OKR quý theo pillar
    - OKR quý link OKR năm
    - OKR tháng (target tháng) theo pillar, link OKR quý
    - Kế hoạch tháng theo OKR tháng
    - KPI con tháng (target + actual) theo OKR tháng
  - Xóa từng bản ghi ngay trên bảng

## Flow liên kết dữ liệu hiện tại

- `OKR năm -> OKR quý -> OKR tháng -> KPI tháng`
- Kế hoạch quý đã được loại khỏi trang nhập liệu.
- Khối Import/Export JSON đã được loại khỏi trang nhập liệu.
- Partner care page:
  - Partner-level summary and touchpoint log
  - Insight and action guidance
- International problem record page:
  - Add case form, SLA metrics, root-cause-oriented analysis

## Notes

- Data updates are stored in browser `localStorage`.
- `Reset theo dữ liệu gốc` reloads from bundled extracted dataset.
- KPI weighting approach uses normalized target contribution by KPI code (monthly target / annual target).
- Optional cloud sharing via Supabase is documented in `SUPABASE_SETUP.md`.
