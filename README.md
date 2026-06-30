# TikTok Affiliate Analyzer

Công cụ phân tích **đơn hàng liên kết / quảng cáo cửa hàng TikTok Shop** từ file Excel export — chạy **hoàn toàn offline** trong trình duyệt, dữ liệu không rời khỏi máy.

Toàn bộ ứng dụng nằm trong **một file HTML duy nhất** (`TikTok_Affiliate_Analyzer.html`, ~1.1 MB) đã nhúng sẵn thư viện đọc Excel và vẽ biểu đồ. Chỉ cần tải về và mở bằng trình duyệt — không cần cài đặt, không cần internet, không upload dữ liệu lên đâu cả.

## Cách dùng

1. Tải file [`TikTok_Affiliate_Analyzer.html`](TikTok_Affiliate_Analyzer.html) về máy.
2. Mở bằng trình duyệt (Chrome / Edge / Firefox).
3. **Kéo–thả** (hoặc bấm chọn) file `.xlsx` export đơn hàng affiliate của TikTok Shop.
4. Xem dashboard, lọc theo cửa hàng / trạng thái / loại nội dung, và đọc phần nhận định + khuyến nghị ở cuối.

## Tính năng

- **6 chỉ số tổng quan**: GMV, hoa hồng ước tính, hoa hồng thực nhận, tỷ lệ quyết toán, tỷ lệ hoàn, hiệu suất hoa hồng.
- **Bộ lọc searchable** (cửa hàng / trạng thái quyết toán / loại nội dung / loại đơn) — mọi biểu đồ & bảng tự tính lại.
- **Phân tích đa chiều**: theo trạng thái quyết toán, theo thời gian (ngày), theo cửa hàng, theo sản phẩm, đơn "không đủ điều kiện", và so sánh ước tính vs thực nhận.
- **Nhận định & khuyến nghị tự sinh**: tóm tắt bức tranh hiện tại và đề xuất hành động phát triển/sản phẩm dựa trên dữ liệu đang lọc.
- **Xuất báo cáo CSV** tổng hợp.
- Giao diện theo phong cách TikTok (nền tối, accent cyan/đỏ), **responsive** cho cả mobile.

## Định dạng file đầu vào

File `.xlsx` export từ TikTok Shop Affiliate (góc nhìn nhà sáng tạo/đối tác), sheet đầu tiên, tiêu đề cột tiếng Việt. Một số cột then chốt: `ID đơn hàng`, `Tên cửa hàng`, `Trạng thái quyết toán đơn hàng`, `Loại nội dung`, `GMV`, `Hoa hồng tiêu chuẩn ước tính`, `Hoa hồng tiêu chuẩn`, `Tổng số tiền nhận được cuối cùng`, `Ngày đặt hàng`.

> Số tiền theo định dạng Việt (`140.000` = 140.000đ) được parse tự động. Cột được dò theo **tên tiếng Việt** nên dùng được cho file export ở các kỳ khác cùng định dạng.

> ⚠️ **Quyền riêng tư**: repo này **không** chứa file dữ liệu thật (`*.xlsx` đã được `.gitignore`). Dữ liệu chỉ được xử lý cục bộ trong trình duyệt của bạn.

## Build từ source

Mã nguồn tách rời nằm trong `build/`:

```
build/
  app.html      # HTML + CSS
  app.js        # Toàn bộ logic (parse, lọc, biểu đồ, nhận định)
  build.mjs     # Script gộp thành 1 file
  lib/          # SheetJS (đọc xlsx) + Chart.js (đã nhúng sẵn)
```

Build lại file đơn:

```bash
node build/build.mjs
```

## Công nghệ

- [SheetJS (xlsx)](https://sheetjs.com/) — đọc file Excel ngay trong trình duyệt.
- [Chart.js](https://www.chartjs.org/) — biểu đồ.
- Vanilla JavaScript, không framework, không backend.
