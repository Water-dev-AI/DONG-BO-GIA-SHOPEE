# Shopee Price Sync 🛒

Công cụ **đồng bộ giá** từ bảng giá mới (CSV / Google Sheet export) vào file Excel chuẩn Shopee, tự động xử lý ràng buộc **giá max/min không vượt quá 5 lần** trong cùng một sản phẩm, làm tròn nghìn, và xuất ra file import sẵn sàng tải lên Shopee Seller.

Chạy được **2 cách**: giao diện web (không cần cài đặt) hoặc dòng lệnh.

---

## ✨ Tính năng

- **Vlookup theo SKU** (`MÃ NỘI BỘ` ↔ `SKU`) để bắt đúng biến thể và cập nhật `GIÁ PEE`.
- **Ràng buộc 5×**: trong cùng một Mã Sản phẩm, `max / min ≤ 5`. Khi vi phạm:
  - Giá thấp → **tăng** (tối đa **+50%**).
  - Giá cao → **giảm** (tối đa **−2%**).
  - Có thể **vừa nâng min vừa hạ max** để hài hoà, ưu tiên thay đổi ít nhất.
- **Làm tròn hàng nghìn** (10.000 ✓, 10.500 ✗).
- **⭕ tự động = max của nhóm** (vì vlookup không ra, đặt max để tránh khách đặt nhầm).
- **Chế độ partial**: nếu CSV chỉ có vài sản phẩm, **xoá giá** các sản phẩm còn lại để Shopee hiểu là không điều chỉnh.
- **An toàn tuyệt đối**: chỉ ghi vào **cột Giá**. Đã kiểm chứng 0 thay đổi ở vùng dữ liệu khác.
- **Tự sửa lỗi file Shopee** (`activePane="bottom_left"`) mà các thư viện Excel thường không đọc được.
- **Báo cáo đầy đủ** (5 sheet): tổng quan, điều chỉnh giá, nhóm cần quyết định, cảnh báo tồn kho, SKU không khớp.
- **Kiểm tra tồn kho**: ⭕ mà số lượng ≠ 0, hoặc SKU thường mà số lượng = 0 → liệt kê để rà soát.

---

## 🖥️ Cách 1 — Giao diện web (khuyên dùng)

Không cần cài Python. Mọi xử lý chạy **ngay trên trình duyệt** (Pyodide), dữ liệu **không** gửi lên server.

### Host miễn phí trên GitHub Pages
1. Fork / push repo này lên GitHub.
2. Vào **Settings → Pages → Build and deployment → Source: GitHub Actions**.
3. Push lên nhánh `main`. Workflow `deploy.yml` sẽ tự build và xuất bản.
4. Truy cập `https://<username>.github.io/<repo>/`.

### Chạy thử cục bộ
```bash
python build_bundle.py          # tạo web/engine_bundle.py
cd web && python -m http.server 8000
# mở http://localhost:8000
```

Thao tác: kéo thả file `.xlsx` Shopee + file `.csv` giá mới → chỉnh tham số (tuỳ chọn) → bấm **Xử lý** → tải về `import_shopee.xlsx` và `bao_cao_dieu_chinh.xlsx`.

---

## ⌨️ Cách 2 — Dòng lệnh (CLI)

```bash
pip install -r requirements.txt

python cli.py \
  --xlsx examples/shopee_mau.xlsx \
  --csv  examples/bang_gia_mau.csv \
  --out  import_shopee.xlsx \
  --report bao_cao.xlsx
```

| Tham số | Mặc định | Ý nghĩa |
|---|---|---|
| `--up` | `0.50` | Tăng tối đa (50% giá bán) |
| `--down` | `0.02` | Giảm tối đa (2% giá bán) |
| `--ratio` | `5` | Giới hạn max/min |
| `--no-clear` | tắt | KHÔNG xoá giá sản phẩm ngoài batch (chế độ full-list) |
| `--keep-lock` | tắt | Giữ nguyên sheet protection |

---

## 📋 Định dạng dữ liệu

**File CSV** (bảng giá mới) cần 2 cột:
- `MÃ NỘI BỘ` — khoá vlookup
- `GIÁ PEE` — giá cần đồng bộ

**File XLSX** (Shopee export):
- Dữ liệu bắt đầu từ **dòng 7** (dòng 1–6 là header Shopee).
- Cột: A `Mã Sản phẩm` · C `Mã Phân loại` · F `SKU` · G `Giá` · I `Số lượng`.

---

## 🧮 Thuật toán ràng buộc 5×

Với mỗi nhóm sản phẩm, nếu `max/min > 5`:

1. Tính biên cho phép mỗi biến thể: hạ tối đa 2%, tăng tối đa 50%.
2. Tìm dải `[L, U]` với `U ≤ L × 5`, sao cho mọi giá kéo được vào dải đó trong biên cho phép.
3. Chọn phương án **chi phí thay đổi nhỏ nhất** (ít xê dịch giá nhất).
4. Làm tròn nghìn và kiểm tra lại tỉ lệ.
5. Nếu không có phương án khả thi → **báo cáo** để bạn tự quyết định.

---

## 🗂️ Cấu trúc dự án

```
shopee-price-sync/
├── core/
│   ├── engine.py        # đọc dữ liệu, vlookup, giải ràng buộc 5×
│   └── writer.py        # ghi file import (chỉ sửa cột giá) + báo cáo
├── web/
│   ├── index.html       # giao diện web
│   └── engine_bundle.py # bundle cho Pyodide (tạo bởi build_bundle.py)
├── examples/            # file mẫu
├── tests/               # bộ test
├── cli.py               # dòng lệnh
├── build_bundle.py      # gộp engine cho web
└── .github/workflows/   # CI + deploy Pages
```

---

## 🧪 Test

```bash
pip install pytest
pytest tests/ -v
```

---

## ⚠️ Lưu ý

- Luôn kiểm tra sheet **"Cần quyết định"** trong báo cáo trước khi import.
- Công cụ mở khoá sheet protection để Shopee import không bị chặn (tắt bằng `--keep-lock`).
- Mọi giá xuất ra đều là số nguyên, tròn nghìn.

## 📄 Giấy phép

MIT
