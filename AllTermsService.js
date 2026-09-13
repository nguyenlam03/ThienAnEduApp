// The aggregate scope is never used as a storage key. Each panel has a real term session.
const ALL_TERMS_CODE_ = '__ALL_TERMS__';

function getAllTerms_() {
  return readObjects_(SHEET_KYHOC)
    .filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED')
    .map(row => ({ maKyHoc: String(row.MaKyHoc || '').trim(), tenKyHoc: String(row.TenKyHoc || '').trim() }))
    .filter(row => row.maKyHoc && row.maKyHoc !== ALL_TERMS_CODE_);
}

function renderAllTermsPage_(page, token, session, brand) {
  const template = HtmlService.createTemplateFromFile('AllTerms');
  template.appUrl = ScriptApp.getService().getUrl();
  template.token = token;
  template.page = page;
  template.brandName = brand.name;
  template.terms = getAllTerms_().map(term => ({
    maKyHoc: term.maKyHoc,
    tenKyHoc: term.tenKyHoc,
    url: template.appUrl + '?page=' + encodeURIComponent(page) + '&embedded=1&token=' +
      encodeURIComponent(SecurityService.createTermSession(token, term.maKyHoc))
  }));
  return template.evaluate().setTitle(brand.name + ' — Tất cả kỳ học')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
