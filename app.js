const { createApp } = Vue;

createApp({
    data() {
        return {
            quy: 'I',
            nam: 2026,
            // Dữ liệu nhập liệu mặc định
            cb1: { dv: null, thu: null, luyKe: null },
            cb2: { dv: null, thu: null, luyKe: null },
            // Tỷ lệ cấu hình theo yêu cầu của PA01 Phú Thọ
            tyLe: { 
                chiBo: 50,           // Trích lại cho Chi bộ 50%
                dangUyGiulLai: 50,   // Đảng ủy giữ lại 50%
                nopCapTren: 30       // Nộp cấp trên 30% của phần Đảng ủy giữ
            }
        }
    },
    computed: {
        // Tổng hợp cơ bản
        tongDv() { return (this.cb1.dv || 0) + (this.cb2.dv || 0); },
        thuKyNay() { return (this.cb1.thu || 0) + (this.cb2.thu || 0); },
        thuLuyKe() { return (this.cb1.luyKe || 0) + (this.cb2.luyKe || 0); },
        
        // TÍNH TOÁN KỲ BÁO CÁO (KỲ NÀY)
        chiBoGiuKyNay() { return Math.round(this.thuKyNay * (this.tyLe.chiBo / 100)); },
        // Nộp cấp trên = (Tổng * 50% Đảng ủy) * 30%
        nopCapTrenKyNay() { return Math.round((this.thuKyNay * (this.tyLe.dangUyGiulLai / 100)) * (this.tyLe.nopCapTren / 100)); },
        // Đảng ủy thực giữ = (Tổng * 50%) - Phần nộp cấp trên
        dangUyGiuKyNay() { return Math.round((this.thuKyNay * (this.tyLe.dangUyGiulLai / 100)) - this.nopCapTrenKyNay); },
        tongGiuLaiKyNay() { return this.chiBoGiuKyNay + this.dangUyGiuKyNay; },

        // TÍNH TOÁN LŨY KẾ (TỪ ĐẦU NĂM)
        chiBoGiuLuyKe() { return Math.round(this.thuLuyKe * (this.tyLe.chiBo / 100)); },
        nopCapTrenLuyKe() { return Math.round((this.thuLuyKe * (this.tyLe.dangUyGiulLai / 100)) * (this.tyLe.nopCapTren / 100)); },
        dangUyGiuLuyKe() { return Math.round((this.thuLuyKe * (this.tyLe.dangUyGiulLai / 100)) - this.nopCapTrenLuyKe); },
        tongGiuLaiLuyKe() { return this.chiBoGiuLuyKe + this.dangUyGiuLuyKe; }
    },
    methods: {
        formatCurrency(value) {
            if (!value) return 0;
            return value.toLocaleString('vi-VN');
        },
        exportExcel() {
            try {
                // 1. Dữ liệu thực tế từ ứng dụng Vue
                const data = {
                    tong_dang_vien: this.tongDv,
                    thu_ky_nay: this.thuKyNay,
                    thu_tu_dau_nam: this.thuLuyKe,
                    giu_ky_nay: this.tongGiuLaiKyNay,
                    cb_giu_ky_nay: this.chiBoGiuKyNay,
                    du_giu_ky_nay: this.dangUyGiuKyNay,
                    giu_luy_ke: this.tongGiuLaiLuyKe,
                    cb_giu_luy_ke: this.chiBoGiuLuyKe,
                    du_giu_luy_ke: this.dangUyGiuLuyKe,
                    nop_ky_nay: this.nopCapTrenKyNay,
                    nop_luy_ke: this.nopCapTrenLuyKe
                };

                // 2. Cấu trúc Mảng 2 chiều (AoA) chuẩn Mẫu B01/ĐP
                const excelData = [
                    ["ĐẢNG CỘNG SẢN VIỆT NAM", "", "", "", "", "", "Mẫu báo cáo B01/ĐP", "", ""],
                    ["Đơn vị báo cáo: Đảng ủy Phòng An ninh đối ngoại", "", "", "", "", "", "Ban hành kèm theo Công văn số 141 -CV/VPTW/nb,", "", ""],
                    ["Đơn vị nhận báo cáo: Văn phòng Đảng ủy Công an tỉnh", "", "", "", "", "", "Ngày 17-3-2011 của Văn phòng Trung ương Đảng", "", ""],
                    [""],
                    ["BÁO CÁO THU NỘP ĐẢNG PHÍ"],
                    [`QUÝ ${this.quy} NĂM ${this.nam}`],
                    ["TT", "Chỉ tiêu", "Đơn vị tính", "Mã số", "Đảng bộ xã, phường, thị trấn", "Đảng bộ doanh nghiệp", "Đảng bộ khác", "Cộng", "Ghi chú"],
                    ["A", "B", "C", "D", "1", "2", "3", "4=1+2+3", "E"],
                    ["I", "Tổng số đảng viên đến cuối kỳ báo cáo", "Người", "01", "", "", data.tong_dang_vien, data.tong_dang_vien, ""],
                    ["II", "Đảng phí đã thu được từ chi bộ của cấp báo cáo", "", "", "", "", "", "", ""],
                    ["1", "Kỳ báo cáo", "Đồng", "02", "", "", data.thu_ky_nay, data.thu_ky_nay, ""],
                    ["2", "Từ đầu năm đến cuối kỳ báo cáo", "Đồng", "03", "", "", data.thu_tu_dau_nam, data.thu_tu_dau_nam, ""],
                    ["III", "Đảng phí trích giữ lại ở các cấp", "", "", "", "", "", "", ""],
                    ["1.0", "Kỳ báo cáo(05+06+07)", "Đồng", "04", "", "", data.giu_ky_nay, data.giu_ky_nay, ""],
                    ["1.1", "Chi bộ", "Đồng", "05", "", "", data.cb_giu_ky_nay, data.cb_giu_ky_nay, ""],
                    ["1.2", "Đảng uỷ cơ sở", "Đồng", "06", "", "", data.du_giu_ky_nay, data.du_giu_ky_nay, ""],
                    ["2.0", "Từ đầu năm đến cuối kỳ báo cáo", "Đồng", "08", "", "", data.giu_luy_ke, data.giu_luy_ke, ""],
                    ["2.1", "Chi bộ", "Đồng", "09", "", "", data.cb_giu_luy_ke, data.cb_giu_luy_ke, ""],
                    ["2.2", "Đảng uỷ cơ sở", "Đồng", "10", "", "", data.du_giu_luy_ke, data.du_giu_luy_ke, ""],
                    ["IV", "Đảng phí nộp lên cấp trên", "", "", "", "", "", "", ""],
                    ["1.0", "Kỳ báo cáo", "Đồng", "12", "", "", data.nop_ky_nay, data.nop_ky_nay, ""],
                    ["2.0", "Từ đầu năm đến cuối kỳ báo cáo", "Đồng", "13", "", "", data.nop_luy_ke, data.nop_luy_ke, ""]
                ];

                // Khởi tạo Worksheet
                const ws = XLSX.utils.aoa_to_sheet(excelData);

                // 3. Cấu hình gộp ô (Merge Cells)
                ws['!merges'] = [
                    { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
                    { s: { r: 0, c: 6 }, e: { r: 0, c: 8 } },
                    { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
                    { s: { r: 1, c: 6 }, e: { r: 1, c: 8 } },
                    { s: { r: 2, c: 0 }, e: { r: 2, c: 4 } },
                    { s: { r: 2, c: 6 }, e: { r: 2, c: 8 } },
                    { s: { r: 4, c: 0 }, e: { r: 4, c: 8 } },
                    { s: { r: 5, c: 0 }, e: { r: 5, c: 8 } },
                ];

                // 4. Định dạng độ rộng cột
                ws['!cols'] = [
                    {wch: 5},   // A: TT
                    {wch: 40},  // B: Chỉ tiêu
                    {wch: 12},  // C: ĐVT
                    {wch: 8},   // D: Mã số
                    {wch: 15},  // E: Xã phường
                    {wch: 15},  // F: Doanh nghiệp
                    {wch: 20},  // G: Đảng bộ khác
                    {wch: 20},  // H: Cộng
                    {wch: 15}   // I: Ghi chú
                ];

                // --- TỰ ĐỘNG STYLE ĐỊNH DẠNG EXCEL CHUẨN THỂ THỨC ---
                const range = XLSX.utils.decode_range(ws['!ref']);
                for (let R = range.s.r; R <= range.e.r; ++R) {
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const cellRef = XLSX.utils.encode_cell({c: C, r: R});
                        if (!ws[cellRef]) continue;

                        // Style mặc định chung: Times New Roman, cỡ 14, tự động xuống dòng (Wrap text)
                        let cellStyle = {
                            font: { name: "Times New Roman", sz: 14, color: { rgb: "000000" } },
                            alignment: { wrapText: true, vertical: "center" }
                        };

                        // Căn giữa và Bôi đậm cho các dòng Tiêu đề trên cùng (Từ dòng 1 đến dòng 8 - theo index từ 0 là R <= 7)
                        if (R <= 7) {
                            cellStyle.font.bold = true;
                            cellStyle.alignment.horizontal = "center";
                        } 
                        
                        // Kẻ khung (Border) mỏng cho toàn bộ bảng biểu (Từ phần tiêu đề cột - index R >= 6 trở đi)
                        if (R >= 6) {
                            cellStyle.border = {
                                top: { style: "thin", color: { auto: 1 } },
                                bottom: { style: "thin", color: { auto: 1 } },
                                left: { style: "thin", color: { auto: 1 } },
                                right: { style: "thin", color: { auto: 1 } }
                            };
                        }

                        // Áp dụng style vào ô
                        ws[cellRef].s = cellStyle;
                    }
                }

                // 5. Tạo Workbook và lưu file
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Bao_Cao_Dang_Phi");
                XLSX.writeFile(wb, `Bao_Cao_B01_DP_PA01_Quy_${this.quy}_Nam_${this.nam}.xlsx`);

            } catch (error) {
                console.error("Lỗi khi xuất file Excel:", error);
                alert("Đã xảy ra lỗi trong quá trình tạo file báo cáo. Vui lòng thử lại!");
            }
        }
    }
}).mount('#app')
