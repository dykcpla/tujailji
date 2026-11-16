/* 엑셀 해석을 화면과 분리된 작업 공간에서 처리 → 앱이 멈추지 않음 */
importScripts('xlsx.mini.min.js', 'excelimport.js');
self.onmessage = function (e) {
  try {
    var s = e.data.b64.replace(/-/g, '+').replace(/_/g, '/'), bin = atob(s), u = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    var o = XIMP.parse(XLSX, u, e.data.prev);
    self.postMessage({ id: e.data.id, ok: true, o: o });
  } catch (err) { self.postMessage({ id: e.data.id, ok: false, msg: String(err && err.message || err) }); }
};
