/* =====================================================================
   xlsx-lite.js — APP KHO CÔNG TRƯỜNG LICOGI13FC · v1.0 · 04/09/2026
   Ghi / đọc file Excel (.xlsx) cho PHIẾU ĐỀ NGHỊ VẬT TƯ (mẫu BM 16.04) mà KHÔNG cần thư viện ngoài
   (CDN bị chặn ở nhiều mạng công trường; app phải chạy được offline).
   - xlsxDeNghi(phieu, anhHeader)  → Blob .xlsx đúng bố cục mẫu công ty (header ảnh, bảng 10 cột, 3 ô ký) — mở bằng Excel sửa được.
   - xlsxDocDeNghi(arrayBuffer)    → Promise<{so, du_an, hang_muc, ngay_de_nghi, ngay_can_ve, phong_mua, ghi_chu, dong[]}>
                                     đọc phiếu do người khác lập ngoài app (Excel theo mẫu BM 16.04, kể cả file đã sửa tay).
   Kỹ thuật: .xlsx = zip; ghi bằng phương thức STORE (không nén) — Excel/LibreOffice/Google Sheets mở bình thường;
   đọc hỗ trợ STORE và DEFLATE (dùng DecompressionStream của trình duyệt, Chrome Android ≥ 80, iOS ≥ 16.4).
   Dùng chung cho trình duyệt và Node (kiểm thử): không dùng DOM.
   ===================================================================== */
(function (root) {
  'use strict';
  var enc = new TextEncoder(), dec = new TextDecoder('utf-8');

  /* ---------- CRC32 + ZIP (store) ---------- */
  var CRC_T = (function () { var t = new Uint32Array(256); for (var i = 0; i < 256; i++) { var c = i; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[i] = c >>> 0; } return t; })();
  function crc32(u8) { var c = 0xFFFFFFFF; for (var i = 0; i < u8.length; i++) c = CRC_T[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function dosTime(d) { return ((d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1)) & 0xFFFF; }
  function dosDate(d) { return (((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xFFFF; }
  /** files: [{name, data: Uint8Array|string}] → Uint8Array zip (STORE) */
  function zipStore(files) {
    var now = new Date(), parts = [], central = [], offset = 0;
    files.forEach(function (f) {
      var name = enc.encode(f.name), data = typeof f.data === 'string' ? enc.encode(f.data) : f.data, crc = crc32(data);
      var lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
      lh.setUint16(10, dosTime(now), true); lh.setUint16(12, dosDate(now), true); lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true); lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
      var ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true);
      ch.setUint16(12, dosTime(now), true); ch.setUint16(14, dosDate(now), true); ch.setUint32(16, crc, true);
      ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, name.length, true);
      ch.setUint16(30, 0, true); ch.setUint16(32, 0, true); ch.setUint16(34, 0, true); ch.setUint16(36, 0, true); ch.setUint32(38, 0, true); ch.setUint32(42, offset, true);
      parts.push(new Uint8Array(lh.buffer), name, data); central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + data.length;
    });
    var cdSize = central.reduce(function (s, p) { return s + p.length; }, 0);
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(4, 0, true); end.setUint16(6, 0, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true); end.setUint16(20, 0, true);
    var all = parts.concat(central, [new Uint8Array(end.buffer)]), tong = all.reduce(function (s, p) { return s + p.length; }, 0);
    var out = new Uint8Array(tong), pos = 0; all.forEach(function (p) { out.set(p, pos); pos += p.length; });
    return out;
  }
  /** Đọc zip (STORE / DEFLATE) → Promise<{name: Uint8Array}> — đọc qua central directory */
  function unzip(buf) {
    var u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf), dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    var eocd = -1; for (var i = u8.length - 22; i >= Math.max(0, u8.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) return Promise.reject(new Error('File không phải .xlsx (không thấy thư mục zip)'));
    var n = dv.getUint16(eocd + 10, true), cd = dv.getUint32(eocd + 16, true), out = {}, jobs = [];
    for (var k = 0; k < n; k++) {
      if (dv.getUint32(cd, true) !== 0x02014b50) break;
      var method = dv.getUint16(cd + 10, true), csize = dv.getUint32(cd + 20, true), nlen = dv.getUint16(cd + 28, true), elen = dv.getUint16(cd + 30, true), clen = dv.getUint16(cd + 32, true), lho = dv.getUint32(cd + 42, true);
      var name = dec.decode(u8.subarray(cd + 46, cd + 46 + nlen));
      var lnlen = dv.getUint16(lho + 26, true), lelen = dv.getUint16(lho + 28, true), start = lho + 30 + lnlen + lelen;
      var data = u8.subarray(start, start + csize);
      if (method === 0) out[name] = data;
      else if (method === 8) jobs.push(inflateRaw(data).then((function (nm) { return function (d) { out[nm] = d; }; })(name)));
      else return Promise.reject(new Error('Zip dùng phương thức nén không hỗ trợ: ' + method));
      cd += 46 + nlen + elen + clen;
    }
    return Promise.all(jobs).then(function () { return out; });
  }
  function inflateRaw(data) {
    if (typeof DecompressionStream === 'undefined') return Promise.reject(new Error('Trình duyệt không hỗ trợ giải nén (cần Chrome ≥ 80 / iOS ≥ 16.4)'));
    var ds = new DecompressionStream('deflate-raw'), w = ds.writable.getWriter(); w.write(data); w.close();
    return new Response(ds.readable).arrayBuffer().then(function (b) { return new Uint8Array(b); });
  }

  /* ---------- tiện ích ---------- */
  function xml(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function colChu(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - 1 - m) / 26; } return s; }
  function colSo(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  function boDau(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().trim(); }
  function serialSangNgay(v) { var d = new Date(Math.round((v - 25569) * 86400 * 1000)); return ('0' + d.getUTCDate()).slice(-2) + '/' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '/' + d.getUTCFullYear(); }

  /* ---------- STYLES (chỉ số s= trong sheet) ----------
     0 thường TNR12 · 1 tiêu đề 18 đậm giữa · 2 số phiếu 13 đậm nghiêng giữa · 3 nhãn đậm trái viền · 4 đầu bảng đậm giữa viền wrap
     5 ô chữ trái viền wrap · 6 ô chữ giữa viền wrap · 7 số phải viền #,##0.### · 8 đậm giữa viền · 9 nghiêng giữa không viền
     10 đậm giữa không viền · 11 chữ trái viền wrap (ghi chú) · 12 chữ phải viền · 13 đậm giữa viền trên (ô ký) · 14 chữ giữa viền trái/phải (khoảng ký)
     15 đậm trái không viền · 16 giữa đậm viền dưới/trái/phải (dòng ngày ký) · 17 đậm giữa viền trái/phải (tên người ký) */
  var STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.###"/></numFmts>' +
    '<fonts count="5"><font><sz val="12"/><name val="Times New Roman"/></font><font><b/><sz val="12"/><name val="Times New Roman"/></font><font><b/><sz val="18"/><name val="Times New Roman"/></font><font><b/><i/><sz val="13"/><name val="Times New Roman"/></font><font><i/><sz val="12"/><name val="Times New Roman"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="6"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>' +
    '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom/><diagonal/></border>' +
    '<border><left style="thin"/><right style="thin"/><top/><bottom/><diagonal/></border>' +
    '<border><left style="thin"/><right style="thin"/><top/><bottom style="thin"/><diagonal/></border>' +
    '<border><left/><right/><top/><bottom style="thin"/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="18">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="2" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="4" fillId="0" borderId="3" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="4" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="3" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  /* ---------- GHI PHIẾU ĐỀ NGHỊ ----------
     phieu = { so, du_an, hang_muc, ngay_de_nghi 'dd/mm/yyyy', ngay_can_ve 'dd/mm/yyyy', phong_mua 'PHONG_MUA'|'DOI_MUA',
               ghi_chu (chuỗi; tách 2 dòng bằng ; hoặc xuống dòng), nguoi_de_nghi, dong: [{ten_vt, tieu_chuan, spec, don_vi, luy_ke_da_nhap, ton_kho, so_luong_de_xuat, so_luong_dat, ghi_chu}] }
     anhHeader = Uint8Array JPEG (ảnh tiêu đề công ty) hoặc null */
  function xlsxDeNghi(p, anhHeader) {
    var cells = {}, hts = {}, merges = [];
    function s(r, c, v, st) { if (v === null || v === undefined) v = ''; cells[r + ':' + c] = { v: v, s: st }; }
    function n(r, c, v, st) { var x = Number(v); cells[r + ':' + c] = { v: (x || x === 0) && String(v) !== '' ? x : '', s: st, num: true }; }
    function m(a) { merges.push(a); }
    var W = { 1: 9, 2: 32, 3: 15.5, 4: 15.5, 5: 8.5, 6: 11.5, 7: 9.5, 8: 9.5, 9: 9.5, 10: 15 };   // rộng cột A..J — vừa 1 trang A4 dọc
    // 1. tiêu đề ảnh + tên phiếu
    hts[1] = 92; m('A1:J1');
    s(2, 1, 'ĐỀ NGHỊ CẤP VẬT TƯ', 1); hts[2] = 27; m('A2:J2');
    s(3, 1, 'Số: ' + (p.so || ''), 2); hts[3] = 22; m('A3:J3');
    s(4, 1, 'DỰ ÁN', 8); s(4, 2, p.du_an || '', 3); s(4, 6, 'NGÀY ĐỀ NGHỊ', 3); s(4, 9, p.ngay_de_nghi || '', 8); hts[4] = 30; m('B4:E4'); m('F4:H4'); m('I4:J4');
    s(5, 1, 'H.MỤC', 8); s(5, 2, p.hang_muc || '', 3); s(5, 6, 'NGÀY VỀ CÔNG TRƯỜNG', 3); s(5, 9, p.ngay_can_ve || '', 8); hts[5] = 24; m('B5:E5'); m('F5:H5'); m('I5:J5');
    // 2. đầu bảng
    var dau = ['STT', 'Tên vật tư', 'Tiêu chuẩn kỹ thuật', 'Spec/Thương hiệu', 'Đơn vị tính', 'Khối lượng', '', '', '', 'Ghi chú'];
    dau.forEach(function (t, i) { s(6, i + 1, t, 4); s(7, i + 1, '', 4); });
    s(7, 6, 'Lũy kế đã nhập hàng', 4); s(7, 7, 'Tồn kho', 4); s(7, 8, 'Đề xuất', 4); s(7, 9, 'Đặt hàng', 4);
    hts[6] = 20; hts[7] = 44; ['A', 'B', 'C', 'D', 'E', 'J'].forEach(function (c) { m(c + '6:' + c + '7'); }); m('F6:I6');
    // 3. dòng vật tư — tối thiểu 12 dòng trống cho giống mẫu
    var dong = p.dong || [], soDong = Math.max(dong.length, 12), r = 8;
    for (var i = 0; i < soDong; i++, r++) {
      var d = dong[i] || {};
      hts[r] = 26;
      s(r, 1, dong[i] ? i + 1 : '', 6); s(r, 2, d.ten_vt || '', 5); s(r, 3, d.tieu_chuan || '', 6); s(r, 4, d.spec || '', 6); s(r, 5, d.don_vi || '', 6);
      n(r, 6, dong[i] ? d.luy_ke_da_nhap : '', 7); n(r, 7, dong[i] ? d.ton_kho : '', 7); n(r, 8, dong[i] ? d.so_luong_de_xuat : '', 7); n(r, 9, dong[i] ? d.so_luong_dat : '', 7); s(r, 10, d.ghi_chu || '', 11);
    }
    s(r, 2, 'Tổng cộng', 8); for (var c = 1; c <= 10; c++) if (c !== 2) s(r, c, '', 6); hts[r] = 22; r++;
    // 4. Phòng mua / Đội mua + ghi chú (2 dòng)
    var gc = String(p.ghi_chu || '').split(/[;\n]/).map(function (x) { return x.trim(); }).filter(Boolean);
    var gc1 = gc[0] ? (/^ghi ch/i.test(gc[0]) ? gc[0] : 'Ghi chú: ' + gc[0]) : 'Ghi chú: Hàng hoá cấp đầy đủ giấy tờ về chất lượng';
    var gc2 = gc[1] ? (/^ghi ch/i.test(gc[1]) ? gc[1] : 'Ghi chú: ' + gc[1]) : 'Ghi chú: ';
    s(r, 1, 'PHÒNG MUA', 12); s(r, 3, p.phong_mua === 'DOI_MUA' ? '' : 'x', 6); s(r, 4, '', 6); s(r, 5, gc1, 11); hts[r] = 22; m('A' + r + ':B' + r); m('E' + r + ':J' + r); var rPm = r; r++;
    s(r, 1, 'ĐỘI MUA', 12); s(r, 3, p.phong_mua === 'DOI_MUA' ? 'x' : '', 6); s(r, 4, '', 6); s(r, 5, gc2, 11); hts[r] = 22; m('A' + r + ':B' + r); m('E' + r + ':J' + r); r++;
    // 5. ô ký
    var rKy = r;
    s(r, 1, 'NGƯỜI ĐỀ NGHỊ', 13); s(r, 3, 'PHỤ TRÁCH KIỂM SOÁT', 13); s(r, 7, 'BAN TỔNG GIÁM ĐỐC', 13);
    [2, 4, 5, 6, 8, 9, 10].forEach(function (c) { s(r, c, '', 13); }); hts[r] = 30; m('A' + r + ':B' + r); m('C' + r + ':F' + r); m('G' + r + ':J' + r); r++;
    for (var k = 0; k < 3; k++, r++) { for (c = 1; c <= 10; c++) s(r, c, '', 14); hts[r] = 22; }
    s(r - 3, 1, p.nguoi_de_nghi || '', 17); m('A' + (r - 3) + ':B' + (r - 1)); m('C' + (r - 3) + ':F' + (r - 1)); m('G' + (r - 3) + ':J' + (r - 1));
    var nd = String(p.ngay_de_nghi || '').split('/');
    s(r, 1, 'Ngày ' + (nd[0] || '……') + ' tháng ' + (nd[1] || '……') + ' năm ' + (nd[2] || '20……'), 16); s(r, 3, 'Ngày …… tháng …… năm 20……', 16); s(r, 7, 'Ngày …… tháng …… năm 20……', 16);
    [2, 4, 5, 6, 8, 9, 10].forEach(function (c) { s(r, c, '', 16); }); hts[r] = 22; m('A' + r + ':B' + r); m('C' + r + ':F' + r); m('G' + r + ':J' + r); r++;
    hts[r] = 8; r++;
    s(r, 1, 'Bộ phận thu mua xác nhận đơn đặt hàng :', 15); hts[r] = 22; m('A' + r + ':J' + r);
    var rCuoi = r;
    // ---- sheet xml
    var rows = [];
    for (var rr = 1; rr <= rCuoi; rr++) {
      var cs = [];
      for (c = 1; c <= 10; c++) {
        var ce = cells[rr + ':' + c]; if (!ce) continue;
        var ref = colChu(c) + rr;
        if (ce.num && ce.v !== '') cs.push('<c r="' + ref + '" s="' + ce.s + '"><v>' + ce.v + '</v></c>');
        else if (ce.v === '' || ce.v === null) cs.push('<c r="' + ref + '" s="' + ce.s + '"/>');
        else cs.push('<c r="' + ref + '" s="' + ce.s + '" t="inlineStr"><is><t xml:space="preserve">' + xml(ce.v) + '</t></is></c>');
      }
      rows.push('<row r="' + rr + '"' + (hts[rr] ? ' ht="' + hts[rr] + '" customHeight="1"' : '') + '>' + cs.join('') + '</row>');
    }
    var cols = '<cols>' + Object.keys(W).map(function (k) { return '<col min="' + k + '" max="' + k + '" width="' + W[k] + '" customWidth="1"/>'; }).join('') + '</cols>';
    var sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:J' + rCuoi + '"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' + cols +
      '<sheetData>' + rows.join('') + '</sheetData><mergeCells count="' + merges.length + '">' + merges.map(function (a) { return '<mergeCell ref="' + a + '"/>'; }).join('') + '</mergeCells>' +
      '<printOptions horizontalCentered="1"/><pageMargins left="0.5" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0"/>' +
      (anhHeader ? '<drawing r:id="rId1"/>' : '') + '</worksheet>';
    var files = [
      { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + (anhHeader ? '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' : '') + '</Types>' },
      { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="DNVT" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">DNVT!$A$1:$J$' + rCuoi + '</definedName></definedNames></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { name: 'xl/styles.xml', data: STYLES },
      { name: 'xl/worksheets/sheet1.xml', data: sheet }
    ];
    if (anhHeader) {
      files.push({ name: 'xl/worksheets/_rels/sheet1.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>' });
      files.push({ name: 'xl/drawings/drawing1.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Header LICOGI13FC"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>' });
      files.push({ name: 'xl/drawings/_rels/drawing1.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/></Relationships>' });
      files.push({ name: 'xl/media/image1.jpeg', data: anhHeader });
    }
    return zipStore(files);
  }

  /* ---------- ĐỌC PHIẾU ĐỀ NGHỊ TỪ FILE EXCEL ---------- */
  function docSheetXml(xmlText, ss) {
    var cells = {}, re = /<c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, m;
    while ((m = re.exec(xmlText))) {
      var col = m[1], row = +m[2], attrs = m[3], inner = m[4] || '', t = (attrs.match(/\bt="(\w+)"/) || [])[1], v = null;
      if (t === 's') { var vi = (inner.match(/<v>([^<]*)<\/v>/) || [])[1]; v = ss[+vi] != null ? ss[+vi] : ''; }
      else if (t === 'inlineStr') v = (inner.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || []).map(function (x) { return x.replace(/<[^>]+>/g, ''); }).join('');
      else if (t === 'str' || t === 'b' || t === 'e') v = (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1] || '';
      else { var vv = (inner.match(/<v>([^<]*)<\/v>/) || [])[1]; v = vv == null || vv === '' ? '' : Number(vv); }
      if (typeof v === 'string') v = giaiXml(v);
      cells[col + row] = v;
      var rc = cells['_r' + row] = cells['_r' + row] || {}; rc[col] = v;
    }
    return cells;
  }
  function giaiXml(s) { return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&'); }
  function docSharedStrings(x) {
    if (!x) return [];
    var out = [], re = /<si>([\s\S]*?)<\/si>/g, m;
    while ((m = re.exec(x))) out.push(giaiXml((m[1].match(/<t[^>]*>([\s\S]*?)<\/t>/g) || []).map(function (t) { return t.replace(/<[^>]+>/g, ''); }).join('')));
    return out;
  }
  function timO(cells, maxRow, dk) {                       // tìm ô đầu tiên thỏa điều kiện dk(text) → {col,row}
    for (var r = 1; r <= maxRow; r++) { var rc = cells['_r' + r]; if (!rc) continue; var ks = Object.keys(rc).sort(function (a, b) { return colSo(a) - colSo(b); }); for (var i = 0; i < ks.length; i++) if (typeof rc[ks[i]] === 'string' && dk(boDau(rc[ks[i]]))) return { col: ks[i], row: r }; }
    return null;
  }
  function oBenPhai(cells, o, boQua) {                     // giá trị đầu tiên khác rỗng bên phải nhãn (bỏ qua các cột trong boQua)
    var rc = cells['_r' + o.row] || {}, ks = Object.keys(rc).filter(function (k) { return colSo(k) > colSo(o.col); }).sort(function (a, b) { return colSo(a) - colSo(b); });
    for (var i = 0; i < ks.length; i++) { var v = rc[ks[i]]; if (v !== '' && v != null && !(boQua && boQua(ks[i], v))) return { col: ks[i], v: v }; }
    return null;
  }
  function chuoiNgay(v) { if (typeof v === 'number') return v > 20000 ? serialSangNgay(v) : String(v); var m = String(v || '').match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/); return m ? ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[3] : String(v || '').trim(); }
  function so(v) { if (typeof v === 'number') return v; var s = String(v || '').replace(/\./g, '').replace(',', '.').trim(); var x = parseFloat(s); return isNaN(x) ? 0 : x; }
  /** arrayBuffer → Promise<phieu>; lỗi rõ nếu không phải mẫu BM 16.04 */
  function xlsxDocDeNghi(buf) {
    return unzip(buf).then(function (z) {
      var wbX = z['xl/workbook.xml'] ? dec.decode(z['xl/workbook.xml']) : '';
      if (!wbX) throw new Error('File không phải Excel .xlsx (thiếu xl/workbook.xml). Nếu là .xls cũ, mở bằng Excel → Lưu dưới dạng .xlsx');
      var ss = docSharedStrings(z['xl/sharedStrings.xml'] ? dec.decode(z['xl/sharedStrings.xml']) : '');
      // chọn sheet: sheet đầu có chữ "ĐỀ NGHỊ CẤP VẬT TƯ", không thì sheet 1
      var sheets = Object.keys(z).filter(function (k) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(k); }).sort(function (a, b) { return (+a.match(/(\d+)\.xml/)[1]) - (+b.match(/(\d+)\.xml/)[1]); });
      if (!sheets.length) throw new Error('File không có sheet nào');
      var chon = null, cells = null;
      for (var i = 0; i < sheets.length; i++) { var cs = docSheetXml(dec.decode(z[sheets[i]]), ss); if (timO(cs, 60, function (t) { return t.indexOf('de nghi cap vat tu') >= 0; })) { chon = sheets[i]; cells = cs; break; } }
      if (!chon) { chon = sheets[0]; cells = docSheetXml(dec.decode(z[chon]), ss); }
      var maxRow = 0; Object.keys(cells).forEach(function (k) { var m = k.match(/^_r(\d+)$/); if (m && +m[1] > maxRow) maxRow = +m[1]; });
      var p = { so: '', du_an: '', hang_muc: '', ngay_de_nghi: '', ngay_can_ve: '', phong_mua: 'PHONG_MUA', ghi_chu: '', nguoi_de_nghi: '', dong: [], canh_bao: [] };
      var oSo = timO(cells, 12, function (t) { return /^so\s*:/.test(t) || /^so\s*$/.test(t); });
      if (oSo) { var vSo = String(cells[oSo.col + oSo.row]).replace(/^s[oố]\s*:?\s*/i, '').trim(); if (!vSo) { var ph = oBenPhai(cells, oSo); vSo = ph ? String(ph.v).trim() : ''; } p.so = vSo; }
      var oDa = timO(cells, 12, function (t) { return t === 'du an' || t.indexOf('du an') === 0; }); if (oDa) { var x = oBenPhai(cells, oDa, function (c, v) { return typeof v === 'string' && /ngay de nghi/.test(boDau(v)); }); p.du_an = x ? String(x.v).trim() : ''; }
      var oHm = timO(cells, 12, function (t) { return /^h\.?\s*muc/.test(t) || t === 'hang muc'; }); if (oHm) { x = oBenPhai(cells, oHm, function (c, v) { return typeof v === 'string' && /ngay ve/.test(boDau(v)); }); p.hang_muc = x ? String(x.v).trim() : ''; }
      var oNd = timO(cells, 12, function (t) { return t.indexOf('ngay de nghi') >= 0; }); if (oNd) { x = oBenPhai(cells, oNd); p.ngay_de_nghi = x ? chuoiNgay(x.v) : ''; }
      var oNv = timO(cells, 12, function (t) { return t.indexOf('ngay ve') >= 0; }); if (oNv) { x = oBenPhai(cells, oNv); p.ngay_can_ve = x ? chuoiNgay(x.v) : ''; }
      // bảng vật tư
      var oTen = timO(cells, 20, function (t) { return t.indexOf('ten vat tu') >= 0; });
      if (!oTen) throw new Error('Không thấy cột "Tên vật tư" — file không theo mẫu BM 16.04');
      var rDau = oTen.row, cot = {};
      var nhan = { ten_vt: /ten vat tu/, tieu_chuan: /tieu chuan/, spec: /spec|thuong hieu/, don_vi: /don vi/, luy_ke_da_nhap: /luy ke/, ton_kho: /ton kho/, so_luong_de_xuat: /de xuat/, so_luong_dat: /dat hang/, ghi_chu: /ghi chu/, stt: /^stt/ };
      [rDau, rDau + 1].forEach(function (r) { var rc = cells['_r' + r] || {}; Object.keys(rc).forEach(function (c) { if (typeof rc[c] !== 'string') return; var t = boDau(rc[c]); Object.keys(nhan).forEach(function (k) { if (!cot[k] && nhan[k].test(t)) cot[k] = c; }); }); });
      if (!cot.so_luong_dat && !cot.so_luong_de_xuat) throw new Error('Không thấy cột "Đặt hàng" / "Đề xuất" trong bảng vật tư');
      var rBd = cells['_r' + (rDau + 1)] && cot.luy_ke_da_nhap && (cells['_r' + (rDau + 1)][cot.luy_ke_da_nhap] != null) ? rDau + 2 : rDau + 1, trong = 0;
      for (var r = rBd; r <= maxRow && trong < 4; r++) {
        var rc = cells['_r' + r] || {}, ten = rc[cot.ten_vt];
        var a = typeof rc.A === 'string' ? boDau(rc.A) : '', b = typeof ten === 'string' ? boDau(ten) : '';
        if (b === 'tong cong' || a === 'tong cong' || /^phong mua|^doi mua|^nguoi de nghi/.test(a) || /^phong mua|^doi mua/.test(b)) break;
        if (ten == null || String(ten).trim() === '') { trong++; continue; }
        trong = 0;
        var d = { ten_vt: String(ten).trim(), tieu_chuan: cot.tieu_chuan ? String(rc[cot.tieu_chuan] == null ? '' : rc[cot.tieu_chuan]).trim() : '', spec: cot.spec ? String(rc[cot.spec] == null ? '' : rc[cot.spec]).trim() : '', don_vi: cot.don_vi ? String(rc[cot.don_vi] == null ? '' : rc[cot.don_vi]).trim() : '',
          luy_ke_da_nhap: cot.luy_ke_da_nhap ? so(rc[cot.luy_ke_da_nhap]) : 0, ton_kho: cot.ton_kho ? so(rc[cot.ton_kho]) : 0, so_luong_de_xuat: cot.so_luong_de_xuat ? so(rc[cot.so_luong_de_xuat]) : 0, so_luong_dat: cot.so_luong_dat ? so(rc[cot.so_luong_dat]) : 0, ghi_chu: cot.ghi_chu ? String(rc[cot.ghi_chu] == null ? '' : rc[cot.ghi_chu]).trim() : '', hang_excel: r };
        if (!d.so_luong_dat && d.so_luong_de_xuat) d.so_luong_dat = d.so_luong_de_xuat;
        if (!d.so_luong_de_xuat && d.so_luong_dat) d.so_luong_de_xuat = d.so_luong_dat;
        p.dong.push(d);
      }
      // phòng mua / đội mua / ghi chú / người đề nghị
      var oPm = timO(cells, maxRow, function (t) { return /^phong mua/.test(t); }), oDm = timO(cells, maxRow, function (t) { return /^doi mua/.test(t); }), gcs = [];
      function tick(o) { if (!o) return false; var rc = cells['_r' + o.row] || {}; return Object.keys(rc).some(function (c) { return colSo(c) > colSo(o.col) && typeof rc[c] === 'string' && /^\s*[xX✓✔v]\s*$/.test(rc[c]); }); }
      function ghiChuHang(o) { if (!o) return; var rc = cells['_r' + o.row] || {}; Object.keys(rc).forEach(function (c) { var v = rc[c]; if (colSo(c) > colSo(o.col) && typeof v === 'string' && v.trim().length > 2 && !/^\s*[xX✓✔v]\s*$/.test(v)) gcs.push(v.replace(/^ghi ch[uú]\s*:?\s*/i, '').trim()); }); }
      ghiChuHang(oPm); ghiChuHang(oDm); p.ghi_chu = gcs.filter(Boolean).join('; ');
      if (tick(oDm) && !tick(oPm)) p.phong_mua = 'DOI_MUA';
      var oNg = timO(cells, maxRow, function (t) { return /^nguoi de nghi/.test(t); });
      if (oNg) { for (r = oNg.row + 1; r <= Math.min(maxRow, oNg.row + 5); r++) { var v = (cells['_r' + r] || {})[oNg.col]; if (typeof v === 'string' && v.trim() && !/^ngay/i.test(v.trim())) { p.nguoi_de_nghi = v.trim(); break; } } }
      if (!p.dong.length) p.canh_bao.push('Không đọc được dòng vật tư nào (bảng trống hoặc sai cột)');
      return p;
    });
  }

  var api = { xlsxDeNghi: xlsxDeNghi, xlsxDocDeNghi: xlsxDocDeNghi, _zipStore: zipStore, _unzip: unzip, _boDau: boDau };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.XLSX_LITE = api;
})(typeof self !== 'undefined' ? self : this);
