// Cloudflare Pages Functions: statik dosyalari API gibi sunar.
// Veriler build sirasinda public/veri/ altina yazilir.
function json(o, status) {
  return new Response(JSON.stringify(o), {
    status: status || 200,
    headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' }
  });
}
function slug(tip, n, sehir) {
  return ('istatistik-' + (tip + '||' + n + '|' + sehir).replace(/[^A-Za-z0-9]+/g, '_'));
}
export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const p = url.pathname;
    const get = (path) => env.ASSETS.fetch(new URL(path, url).toString());

    if (p === '/api/gun') {
      const t = url.searchParams.get('tarih') || '';
      let key = '/veri/gun.json';
      const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(t);
      if (m) key = '/veri/gun-' + m[1] + m[2] + m[3] + '.json';
      let r = await get(key);
      if (!r.ok) r = await get('/veri/gun.json');
      return r;
    }
    if (p === '/api/turler') return get('/veri/turler.json');
    if (p === '/api/sehirler') return get('/veri/sehirler.json');
    if (p === '/api/jokeyler') {
      const sid = url.searchParams.get('sehir') || '-1';
      const yil = url.searchParams.get('yil') || String(new Date().getFullYear());
      let r = await get('/veri/jokey-' + sid + '-' + yil + '.json');
      if (!r.ok) r = await get('/veri/jokey-' + sid + '-' + new Date().getFullYear() + '.json');
      if (!r.ok) return json({ SehirId: sid, Yil: yil, Jokeyler: [] });
      return r;
    }
    if (p === '/api/istatistik') {
      const tip = ((url.searchParams.get('tip') || '').split('/')[0]).trim();
      const n = url.searchParams.get('n') || '50';
      const sehir = url.searchParams.get('sehir') || '-1';
      const r = await get('/veri/' + slug(tip, n, sehir) + '.json');
      if (r.ok) return r;
      return json({ Hazirlaniyor: true, Tip: tip, KosuSayisi: 0, Mesaj: 'İstatistik hazırlanıyor; birkaç dakika içinde otomatik gelecek.' });
    }
    if (p === '/api/analiz') {
      const kosu = (url.searchParams.get('kosu') || '').replace(/[^A-Za-z0-9_-]/g, '_');
      const r = await get('/veri/analiz-' + kosu + '.json');
      if (r.ok) return r;
      return json({ error: 'analiz yok' }, 404);
    }
    if (p === '/istatistik') {
      const tip = ((url.searchParams.get('tip') || '').split('/')[0]).trim();
      const n = url.searchParams.get('n') || '50';
      const sehir = url.searchParams.get('sehir') || '-1';
      let r = await get('/' + slug(tip, n, sehir) + '.html');
      if (!r.ok) r = await get('/turler.html');
      return r;
    }
    return env.ASSETS.fetch(request);
  }
};
