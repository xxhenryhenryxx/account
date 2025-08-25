/**
 * Ghi nhiều tài khoản vào sheet lần lượt từ trên xuống dưới
 * @param {Array} danhSachTaiKhoan - Mảng chứa mã tài khoản cần ghi
 * @returns {Object} Kết quả ghi dữ liệu
 */
function ghiNhieuTaiKhoanVaoSheet(danhSachTaiKhoan) {
  try {
    // Kiểm tra dữ liệu đầu vào
    if (!danhSachTaiKhoan || !Array.isArray(danhSachTaiKhoan) || danhSachTaiKhoan.length === 0) {
      return {
        success: false,
        error: 'Danh sách tài khoản không hợp lệ'
      };
    }

    // Lấy sheet hiện tại
    const sheet = SpreadsheetApp.getActiveSheet();
    if (!sheet) {
      return {
        success: false,
        error: 'Không thể lấy sheet hiện tại'
      };
    }

    // Lấy cell đang được chọn (active cell)
    const activeCell = sheet.getActiveCell();
    if (!activeCell) {
      return {
        success: false,
        error: 'Không thể lấy cell đang được chọn'
      };
    }

    // Lấy vị trí bắt đầu (dòng và cột của cell được chọn)
    const startRow = activeCell.getRow();
    const startCol = activeCell.getColumn();
    
    console.log(`Bắt đầu ghi từ dòng ${startRow}, cột ${startCol}`);

    // Ghi từng tài khoản lần lượt từ trên xuống dưới
    let successCount = 0;
    let errorCount = 0;

    danhSachTaiKhoan.forEach((maTaiKhoan, index) => {
      try {
        // Tính toán vị trí dòng để ghi (dòng đầu + index)
        const targetRow = startRow + index;
        
        // Ghi mã tài khoản vào cell
        sheet.getRange(targetRow, startCol).setValue(maTaiKhoan);
        
        console.log(`Đã ghi "${maTaiKhoan}" vào dòng ${targetRow}, cột ${startCol}`);
        successCount++;
        
      } catch (error) {
        console.error(`Lỗi khi ghi tài khoản "${maTaiKhoan}":`, error.toString());
        errorCount++;
      }
    });

    // Trả về kết quả
    if (errorCount === 0) {
      return {
        success: true,
        count: successCount,
        message: `Đã ghi thành công ${successCount} tài khoản`
      };
    } else {
      return {
        success: true,
        count: successCount,
        errorCount: errorCount,
        message: `Đã ghi ${successCount} tài khoản, ${errorCount} lỗi`
      };
    }

  } catch (error) {
    console.error('Lỗi trong hàm ghiNhieuTaiKhoanVaoSheet:', error.toString());
    return {
      success: false,
      error: 'Lỗi hệ thống: ' + error.toString()
    };
  }
}

