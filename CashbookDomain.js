/**
 * Cashbook Domain
 * Chuẩn hóa lệnh ghi sổ trước khi Repository thực hiện ghi dữ liệu.
 */
var CashbookDomain = (function () {
  function text(value, maxLength) {
    return String(value == null ? '' : value).trim().slice(0, maxLength || 1000);
  }

  function validDate(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
    var parts = value.split('-').map(Number);
    var date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    return date.getUTCFullYear() === parts[0] && date.getUTCMonth() === parts[1] - 1 && date.getUTCDate() === parts[2];
  }

  function transaction(input) {
    input = input || {};
    var command = {
      maGiaoDich: text(input.maGiaoDich, 80),
      loai: text(input.loai, 10).toUpperCase(),
      maDanhMuc: text(input.maDanhMuc, 80),
      maKeHoachChi: text(input.maKeHoachChi, 80),
      maNhanSu: text(input.maNhanSu, 80),
      maHuTaiChinh: text(input.maHuTaiChinh, 40).toUpperCase(),
      maNguonTien: text(input.maNguonTien, 40).toUpperCase(),
      ngayGiaoDich: text(input.ngayGiaoDich, 10),
      noiDung: text(input.noiDung, 500),
      soTien: Math.max(0, Number(input.soTien) || 0),
      nguoiNopNhan: text(input.nguoiNopNhan, 200),
      soChungTu: text(input.soChungTu, 100),
      ghiChu: text(input.ghiChu, 1000),
      chungTuImage: input.chungTuImage || null
    };
    if (command.loai !== 'THU' && command.loai !== 'CHI') throw new Error('Loại giao dịch không hợp lệ.');
    if (!validDate(command.ngayGiaoDich)) throw new Error('Vui lòng nhập ngày giao dịch hợp lệ.');
    if (!command.maDanhMuc) throw new Error('Vui lòng chọn danh mục thu chi.');
    if (!command.maNguonTien) throw new Error('Vui lòng chọn nguồn tiền.');
    if (!command.noiDung) throw new Error('Vui lòng nhập nội dung thu chi.');
    if (command.soTien <= 0) throw new Error('Số tiền phải lớn hơn 0.');
    return command;
  }

  function transfer(input) {
    input = input || {};
    var command = {
      ngayGiaoDich: text(input.ngayGiaoDich, 10),
      maNguonDi: text(input.maNguonDi, 40).toUpperCase(),
      maNguonDen: text(input.maNguonDen, 40).toUpperCase(),
      soTien: Math.max(0, Number(input.soTien) || 0),
      noiDung: text(input.noiDung, 500) || 'Chuyển tiền nội bộ',
      soChungTu: text(input.soChungTu, 100),
      ghiChu: text(input.ghiChu, 1000)
    };
    if (!validDate(command.ngayGiaoDich)) throw new Error('Vui lòng nhập ngày chuyển tiền hợp lệ.');
    if (!command.maNguonDi || !command.maNguonDen || command.maNguonDi === command.maNguonDen) {
      throw new Error('Nguồn chuyển và nguồn nhận phải khác nhau.');
    }
    if (command.soTien <= 0) throw new Error('Số tiền chuyển phải lớn hơn 0.');
    return command;
  }

  return { transaction: transaction, transfer: transfer, validDate: validDate };
})();
