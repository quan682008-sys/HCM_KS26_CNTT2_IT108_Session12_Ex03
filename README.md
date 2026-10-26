# HCM_KS26_CNTT2_IT108_Session12_Ex03
##  Phần 1. Phân tích

Hàm VLOOKUP giống như nhân viên kho tìm hồ sơ vận đơn:

| Tham số | Ý nghĩa thực tế |
|---|---|
| lookup_value | Mã vận đơn cần tìm |
| table_array | Bảng hồ sơ chứa mã và giá cước |
| col_index_num | Cột thứ 2 chứa giá cước |
| range_lookup | FALSE để tìm chính xác mã |

---

##  Phần 2. Triển khai

- **Word:** Chuẩn hóa quy trình nhận hàng, kiểm tra, tra cước, phân tuyến và bàn giao.
- **Excel:** Dùng VLOOKUP kết hợp $ để cố định vùng dữ liệu, áp dụng cho 500 vận đơn.
- **PowerPoint:** Trình bày tối đa 4 ý chính mỗi slide, dùng tiêu đề thể hiện kết luận.

**Công thức Excel:**

```excel
=IFERROR(VLOOKUP(TRIM(CLEAN(SUBSTITUTE(A2,CHAR(160)," "))),BangGiaCuoc!$A$2:$B$501,2,FALSE),"KHÔNG TÌM THẤY")
```

---

##  Phần 3. Kiểm thử

| Tình huống | Kết quả mong đợi |
|---|---|
| Mã hợp lệ | Trả về đúng giá cước |
| Mã dư khoảng trắng | TRIM xử lý và tìm được giá |
| Mã chứa ký tự lạ # | Hiện “KHÔNG TÌM THẤY” |
