/* Kalkulator Akustyki - logika okna. Obliczenia, raport i pliki: Python (interfejs.py) przez window.pywebview.api */
'use strict';

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = (t) => String(t ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const lp = (n, m = 2) => Number(n || 0).toLocaleString('pl-PL', { minimumFractionDigits: m, maximumFractionDigits: m });
const lc = (n) => Number(n || 0).toLocaleString('pl-PL', { maximumFractionDigits: 2 });
const nazwaPliku = (p) => String(p || '').split(/[\\/]/).pop();

const IK = {
  ok: '<path d="M20 6 9 17l-5-5"/>',
  x: '<path d="M18 6 6 18M6 6l12 12"/>',
  plik: '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8l-5-5Z"/><path d="M14 3v5h5M9 13h6M9 17h4"/>',
  pobierz: '<path d="M12 3v12M7 10l5 5 5-5M5 21h14"/>',
  odswiez: '<path d="M21 12a9 9 0 1 1-3-6.7L21 8M21 3v5h-5"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  kosz: '<path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6"/>',
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 8h.01M11 12h1v5h1"/>',
  uwaga: '<path d="M12 9v4M12 17h.01M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/>',
  folder: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"/>',
  zapisz: '<path d="M5 3h11l5 5v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z"/><path d="M7 3v5h8M7 21v-7h10v7"/>',
  nowy: '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8l-5-5Z"/><path d="M14 3v5h5M12 11v6M9 14h6"/>',
  otworz: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v1M3 7v10a2 2 0 0 0 2 2h12l4-8H7l-4 8"/>',
  word: '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8l-5-5Z"/><path d="M14 3v5h5M8 12l1.3 5L11 13l1.7 4L14 12"/>',
  piorun: '<path d="M13 2 4 14h7l-1 8 9-12h-7l1-8Z"/>',
  bateria: '<rect x="2" y="7" width="18" height="11" rx="2"/><path d="M22 11v3M6 11v3M10 11v3"/>',
  kabel: '<path d="M4 4v5a4 4 0 0 0 4 4h8a4 4 0 0 1 4 4v3M2 4h4M18 20h4"/>',
  wtyczka: '<path d="M9 2v6M15 2v6M6 8h12v4a6 6 0 0 1-12 0V8ZM12 18v4"/>',
  zegar: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  silnik: '<rect x="3" y="7" width="18" height="13" rx="2"/><path d="M7 7V4h3v3M14 7V4h3v3M13 10l-3 4h4l-3 4"/>',
  obraz: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="m21 16-5-5-9 9"/>',
  dokument: '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8l-5-5Z"/><path d="M14 3v5h5"/>',
  dalej: '<path d="M9 6l6 6-6 6"/>',
  wyslij: '<path d="M22 2 11 13M22 2l-7 20-4-9-9-4Z"/>',
  baza: '<ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3"/>',
  szukaj: '<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>',
  warstwy: '<path d="m12 3 9 5-9 5-9-5 9-5Z"/><path d="m3 13 9 5 9-5M3 17l9 5 9-5"/>',
  obudowa: '<path d="M3 8l9-5 9 5v8l-9 5-9-5Z"/><path d="M3 8l9 5 9-5M12 13v8"/>',
  wiatrak: '<circle cx="12" cy="12" r="2"/><path d="M12 10c0-4 1-7 4-7s2 5-2 8M14 12c4 0 7 1 7 4s-5 2-8-2M12 14c0 4-1 7-4 7s-2-5 2-8M10 12c-4 0-7-1-7-4s5-2 8 2"/>',
};
const ik = (n, kl = '') => `<svg class="ik ${kl}" viewBox="0 0 24 24">${IK[n]}</svg>`;

// ------------------------------------------------------------------ stan okna
let INFO = {}, BAZA = { materials: {}, engines: {}, radiators: {} }, P = null, W = null, AKT = null, AKT_BAZY = null, KONTO = { zalogowany: false };
let ZAPISANY = '';      // projekt w chwili ostatniego zapisu / wczytania (do wykrywania zmian)
let widok = 'obliczenia';
let przeliczTimer = null;
const baza = { rodzaj: 'materials', wybrany: null, szukaj: '' };
const MINIATURY = {};  // ścieżka -> data URI (albo '' gdy pliku nie ma)

async function api(nazwa, ...arg) {
  try {
    return await window.pywebview.api[nazwa](...arg);
  } catch (e) {
    toast(`Błąd programu (${nazwa}): ${esc(e.message || e)}`, 'blad');
    console.error(e);
    return null;
  }
}

// ------------------------------------------------------------------ motyw i rozmiar
const ciemnySystem = window.matchMedia('(prefers-color-scheme: dark)');
function ustawMotyw(m) {
  document.documentElement.dataset.motyw = m === 'system' ? (ciemnySystem.matches ? 'ciemny' : 'jasny') : m;
  $$('#motyw button').forEach((b) => b.classList.toggle('akt', b.dataset.motyw === m));
}
ciemnySystem.addEventListener('change', () => INFO.program && ustawMotyw(INFO.program.motyw));
async function zmienProgram(klucz, wartosc) {
  INFO.program = await api('ustaw_program', klucz, wartosc);
  if (klucz === 'motyw') ustawMotyw(wartosc);
  if (klucz === 'rozmiar') document.documentElement.dataset.rozmiar = wartosc;
}

// ------------------------------------------------------------------ powiadomienia i okna dialogowe
function toast(tekst, rodzaj = '', akcje = null, czas = 5000) {
  const el = document.createElement('div');
  el.className = `toast ${rodzaj}`;
  el.innerHTML = `${ik(rodzaj === 'blad' ? 'uwaga' : rodzaj === 'ok' ? 'ok' : 'info')}<span>${tekst}</span>`;
  for (const a of [].concat(akcje || [])) {
    const b = document.createElement('button');
    b.className = 'btn btn-maly';
    b.textContent = a.tekst;
    b.onclick = () => { a.fn(); el.remove(); };
    el.append(b);
  }
  $('#powiadomienia').append(el);
  setTimeout(() => el.remove(), rodzaj === 'blad' ? czas * 2 : czas);
}

let modalOtwarty = null;
function modal({ tytul, tresc = '', przyciski = [], szeroki = false, poOtwarciu = null }) {
  return new Promise((gotowe) => {
    const tlo = document.createElement('div');
    tlo.className = 'modal-tlo';
    tlo.innerHTML = `<div class="modal ${szeroki ? 'szeroki' : ''}" role="dialog"><h2>${tytul}</h2><div class="tresc-modala">${tresc}</div>
      ${przyciski.length ? `<div class="przyciski">${przyciski.map((p, i) => `<button class="btn ${p.klasa || ''}" data-i="${i}">${p.tekst}</button>`).join('')}</div>` : ''}</div>`;
    const zamknij = (w) => { tlo.remove(); modalOtwarty = null; gotowe(w); };
    tlo.addEventListener('click', (e) => {
      if (e.target === tlo) zamknij(null);
      const b = e.target.closest('[data-i]');
      if (b) zamknij(przyciski[+b.dataset.i].wartosc);
      const w = e.target.closest('[data-wybor]');
      if (w) zamknij(w.dataset.wybor);
    });
    modalOtwarty = { zamknij, el: tlo };
    $('#modal-miejsce').append(tlo);
    const dom = przyciski.findIndex((p) => p.domyslny);
    (dom >= 0 ? $(`[data-i="${dom}"]`, tlo) : $('input, textarea, .btn', tlo))?.focus();
    poOtwarciu?.(tlo, zamknij);
  });
}

// ------------------------------------------------------------------ start
window.addEventListener('pywebviewready', start);

async function start() {
  INFO = await api('informacje');
  ustawMotyw(INFO.program.motyw);
  document.documentElement.dataset.rozmiar = INFO.program.rozmiar;
  $('#start-wersja').textContent = $('#boczne-wersja').textContent = `DIAMEL · v${INFO.wersja}`;
  BAZA = INFO.baza;
  KONTO = INFO.konto;
  ustawProjekt(INFO.projekt);
  pokaz('obliczenia');
  if (INFO.przeniesiono_baze) toast('Twoje własne materiały, silniki i chłodnice z poprzedniej wersji programu zostały przeniesione do nowej bazy.', 'ok', null, 9000);
  sprawdzAktualizacje(false);
  sprawdzBaze(false);
  if (KONTO.zalogowany) odswiezKonto(); else ekranLogowania();
}

// Logowanie przy starcie (jak w Korektorze HR). Konto nie jest obowiązkowe - "Pracuj bez logowania".
function ekranLogowania(blad = '') {
  $('#ekran-start').classList.remove('ukryty');
  const t = $('#start-tresc');
  t.innerHTML = `
    <div><h2 style="font-size:1.3rem">Zaloguj się</h2><div class="podp">Konto Diamel: papier firmowy Twojej firmy w raportach, inicjały i karty producentów z bazy (silniki, chłodnice, materiały)</div></div>
    ${blad ? `<div class="blad">${esc(blad)}</div>` : ''}
    <div><label class="etykieta" for="l-login">Login</label><input class="pole" id="l-login" autocomplete="username" spellcheck="false"></div>
    <div><label class="etykieta" for="l-haslo">Hasło</label><input class="pole" id="l-haslo" type="password" autocomplete="current-password"></div>
    <label class="zaznacz"><input type="checkbox" id="l-zapamietaj" checked> Zapamiętaj na tym komputerze</label>
    <button class="btn btn-hr duzy" id="l-zaloguj">Zaloguj</button>
    <div style="display:flex; justify-content:space-between; gap:10px; font-size:0.86rem">
      <button class="link" id="l-przypomnij">Nie pamiętam hasła</button>
    </div>
    <button class="btn btn-bez btn-maly" id="l-bez">Pracuj bez logowania</button>`;
  const zaloguj = async () => {
    const login = $('#l-login').value.trim(), haslo = $('#l-haslo').value;
    if (!login || !haslo) return $(login ? '#l-haslo' : '#l-login').focus();
    $('#l-zaloguj').disabled = true;
    $('#l-zaloguj').textContent = 'Logowanie…';
    const r = await api('konto_zaloguj', login, haslo, $('#l-zapamietaj').checked);
    if (!r?.ok) return ekranLogowania(r?.blad || 'Logowanie nie powiodło się.');
    KONTO = r.stan;
    $('#ekran-start').classList.add('ukryty');
    toast(`Zalogowano: <b>${esc(KONTO.nazwa)}</b>${KONTO.szablon ? `<br>Raporty na papierze firmowym ${esc(KONTO.firma)}.` : ''}`, 'ok', null, 6000);
    if (r.uwaga) toast(esc(r.uwaga), 'blad');
    poZmianieKonta();
  };
  $('#l-zaloguj').onclick = zaloguj;
  $$('#l-login, #l-haslo').forEach((el) => el.addEventListener('keydown', (e) => e.key === 'Enter' && zaloguj()));
  $('#l-przypomnij').onclick = async () => {
    const login = $('#l-login').value.trim();
    if (!login) { toast('Wpisz login - wyślemy prośbę o nowe hasło do administratora.'); return $('#l-login').focus(); }
    const r = await api('konto_przypomnij', login);
    if (r?.ok) toast('Prośba o nowe hasło została wysłana. Administrator przyśle nowe hasło e-mailem.', 'ok', null, 9000);
    else toast(esc(r?.blad || 'Nie udało się wysłać prośby.'), 'blad');
  };
  $('#l-bez').onclick = () => $('#ekran-start').classList.add('ukryty');
  $('#l-login').focus();
}

// ------------------------------------------------------------------ konto Diamel (nieobowiązkowe): szablon raportu firmy, inicjały
async function odswiezKonto() {
  if (!KONTO.zalogowany) return;
  const r = await api('konto_odswiez');
  if (!r) return;
  KONTO = r.stan;
  if (r.wylogowano) ekranLogowania(`Wylogowano z konta Diamel: ${r.wylogowano}`);
  odswiezDol();
  if (widok === 'ustawienia') rysujUstawienia();
  if (widok === 'obliczenia') rysujWyniki();
}

/** Po zalogowaniu / wylogowaniu: nowy, nieruszony projekt dostaje inicjały z konta. */
async function poZmianieKonta() {
  if (!P._plik && !czyNiezapisane()) ustawProjekt(await api('nowy_projekt'));
  odswiezDol();
  rysujUstawienia();
}

function opisSzablonu() {
  if (INFO.program.szablon) return `stopka z szablonu ${nazwaPliku(INFO.program.szablon)}`;
  if (KONTO.szablon) return `papier firmowy ${KONTO.firma}`;
  return 'stopka programu';
}

function pokaz(nazwa) {
  widok = nazwa;
  $$('.poz').forEach((b) => b.classList.toggle('akt', b.dataset.widok === nazwa));
  $$('.widok').forEach((s) => s.classList.toggle('ukryty', s.id !== `widok-${nazwa}`));
  if (nazwa === 'obliczenia') rysujObliczenia();
  if (nazwa === 'baza') rysujBaze();
  if (nazwa === 'ustawienia') rysujUstawienia();
  if (nazwa === 'zglos') rysujZglos();
  rysujPasek();
}
$$('.poz').forEach((b) => b.addEventListener('click', () => pokaz(b.dataset.widok)));
$$('#motyw button').forEach((b) => b.addEventListener('click', () => zmienProgram('motyw', b.dataset.motyw)));

function odswiezDol() {
  $('#dol-konto').innerHTML = KONTO.zalogowany ? `Zalogowano: <b>${esc(KONTO.login)}</b>${KONTO.firma ? ` · ${esc(KONTO.firma)}` : ''}` : 'Bez logowania';
  $('#dol-baza').innerHTML = `Baza: <b>v${esc(BAZA.wersja)}</b>${AKT_BAZY?.nowsza ? ` · <span class="akt-nowa" id="dol-baza-akt">nowa v${esc(AKT_BAZY.wersja)}</span>` : ''}`;
  const w = $('#dol-wersja');
  w.innerHTML = AKT?.do_instalacji ? `<span class="akt-nowa">● ${AKT.wycofana ? 'zalecana wersja' : 'dostępna wersja'} ${esc(AKT.wersja)}</span>`
    : `<span class="akt-ok">● v${esc(INFO.wersja || '')}${AKT?.ok ? ' · aktualna' : ''}</span>`;
  w.onclick = AKT?.do_instalacji ? oknoAktualizacji : null;
  $('#dol-baza-akt')?.addEventListener('click', () => aktualizujBaze(false));
  $('#licznik-baza').classList.toggle('ukryty', !AKT_BAZY?.nowsza);
  $('#licznik-baza').textContent = 'nowa';
  $('#licznik-ustawienia').classList.toggle('ukryty', !AKT?.do_instalacji);
  $('#licznik-ustawienia').textContent = '1';
}

// ------------------------------------------------------------------ aktualizacje programu i bazy
async function sprawdzAktualizacje(recznie) {
  AKT = await api('sprawdz_aktualizacje');
  odswiezDol();
  if (widok === 'ustawienia') rysujUstawienia();
  const miejsce = $('#start-aktualizacja');
  if (AKT?.do_instalacji) {
    miejsce.innerHTML = `<div class="baner zolty">${ik('pobierz')}<div class="tekst"><b>${AKT.wycofana ? `Zalecany powrót do wersji ${esc(AKT.wersja)}` : `Dostępna wersja ${esc(AKT.wersja)}`}</b>${AKT.wycofana ? 'W tej wersji wykryto błąd - zainstaluj zalecaną.' : 'Możesz zainstalować ją przed zalogowaniem.'}</div>
      <button class="btn btn-maly" id="start-instaluj">Zainstaluj</button></div>`;
    $('#start-instaluj').onclick = oknoAktualizacji;
  } else miejsce.innerHTML = '';
  if (recznie) {
    if (!AKT?.ok) toast(`Nie udało się sprawdzić aktualizacji: ${esc(AKT?.blad || '')}`, 'blad');
    else if (AKT.do_instalacji) oknoAktualizacji();
    else toast(`Masz najnowszą wersję programu (${esc(INFO.wersja)}).`, 'ok');
  } else if (AKT?.do_instalacji) oknoAktualizacji(); // przy starcie pytanie: pobrać i zainstalować czy później
}

async function oknoAktualizacji() {
  const w = await modal({
    tytul: `${ik('pobierz')} ${AKT.wycofana ? 'Powrót do wersji' : 'Aktualizacja do wersji'} ${esc(AKT.wersja)}`, szeroki: true,
    tresc: `<p>Obecna wersja: <b>${esc(INFO.wersja)}</b>. ${AKT.wycofana ? 'W tej wersji wykryto błąd - zalecana jest poprzednia wersja. Twoje projekty, ustawienia i baza zostaną. ' : ''}Program zamknie się na czas instalacji.</p>
      <pre class="changelog">${esc(AKT.opis || 'Brak opisu zmian.')}</pre>`,
    przyciski: [{ tekst: 'Później', wartosc: null }, { tekst: `${ik('pobierz')} Pobierz i zainstaluj`, klasa: 'btn-hr', wartosc: 'tak', domyslny: true }],
  });
  if (w !== 'tak') return;
  if (czyNiezapisane() && !(await pytajOZapis('Zainstalować aktualizację?'))) return;
  modal({ tytul: 'Pobieranie aktualizacji…', tresc: '<div class="pasek-wyniku"><i id="akt-postep" style="width:0%"></i></div><p class="podp" id="akt-tekst" style="margin-top:10px">0%</p>' });
  window.postepAktualizacji = (p) => { const el = $('#akt-postep'); if (el) { el.style.width = `${p}%`; $('#akt-tekst').textContent = `${p}%`; } };
  window.aktualizacjaPobrana = (ok, blad) => {
    if (ok) {
      const t = $('#akt-tekst');
      if (t) t.textContent = 'Instalowanie… Program zamknie się na chwilę i uruchomi ponownie sam.';
      return setTimeout(() => api('uruchom_instalator'), 900);
    }
    modalOtwarty?.zamknij();
    toast(`Nie udało się pobrać aktualizacji: ${esc(blad)}`, 'blad');
  };
  api('pobierz_aktualizacje', AKT.link);
}

async function sprawdzBaze(recznie) {
  AKT_BAZY = await api('sprawdz_baze');
  odswiezDol();
  if (widok === 'baza') rysujBaze();
  if (recznie) {
    if (!AKT_BAZY?.ok) toast(`Nie udało się sprawdzić bazy: ${esc(AKT_BAZY?.blad || '')}`, 'blad');
    else aktualizujBaze(false);
  } else if (AKT_BAZY?.nowsza) {
    aktualizujBaze(true);  // jak lista pracowników w HR: nowa baza globalna pobierana od razu przy starcie
  }
}

/** Pełna aktualizacja bazy globalnej z serwera - pobrana baza zastępuje poprzednią, wpisy własne zostają. */
async function aktualizujBaze(auto) {
  if (!auto) {
    const wlasne = Object.values(BAZA.zrodla || {}).reduce((n, z) => n + Object.values(z).filter((x) => x !== 'globalna').length, 0);
    const w = await modal({
      tytul: `${ik('baza')} Pobierz bazę z serwera`,
      tresc: `<p>Twoja wersja: <b>v${esc(BAZA.wersja)}</b>${AKT_BAZY?.ok ? `, na serwerze: <b>v${esc(AKT_BAZY.wersja)}</b>` : ''}. Baza globalna Diamel zostanie pobrana w całości i zastąpi obecną - także pozycje usunięte z bazy globalnej znikną z listy.</p>
        <p>${wlasne ? `Twoje wpisy własne i zmienione (${wlasne}) zostaną bez zmian.` : 'Nie masz wpisów własnych.'} Przed zmianą program zrobi kopię obecnej bazy.</p>
        ${AKT_BAZY?.opis ? `<pre class="changelog">${esc(AKT_BAZY.opis)}</pre>` : ''}`,
      przyciski: [{ tekst: 'Anuluj', wartosc: null }, { tekst: `${ik('pobierz')} Pobierz bazę`, klasa: 'btn-hr', wartosc: 'tak', domyslny: true }],
    });
    if (w !== 'tak') return;
  }
  const r = await api('aktualizuj_baze');
  if (!r?.ok) return toast(esc(r?.blad || 'Aktualizacja bazy nie powiodła się.'), 'blad');
  BAZA = r.baza;
  AKT_BAZY = { ...(AKT_BAZY || {}), ok: true, nowsza: false, wersja: r.jest };
  const opis = (rodz, n) => {
    const z = r.zmiany[rodz] || { nowe: [], zmienione: [], usuniete: [] };
    const c = [z.nowe.length && `nowe ${z.nowe.length}`, z.zmienione.length && `zmienione ${z.zmienione.length}`, z.usuniete.length && `usunięte ${z.usuniete.length}`].filter(Boolean);
    return `${n}: ${c.join(', ') || 'bez zmian'}`;
  };
  toast(`Baza ${r.bylo === r.jest ? `pobrana ponownie (v${esc(r.jest)})` : `zaktualizowana: v${esc(r.bylo)} → <b>v${esc(r.jest)}</b>`}<br>${opis('materials', 'Materiały')} · ${opis('engines', 'Silniki')} · ${opis('radiators', 'Chłodnice')}`,
    'ok', { tekst: 'Kopia', fn: () => api('otworz_folder', r.kopia) }, 10000);
  odswiezDol();
  if (widok === 'baza') rysujBaze();
  if (widok === 'obliczenia') rysujObliczenia();
}

// ------------------------------------------------------------------ projekt: nowy, wczytaj, zapisz
const doPorownania = (p) => JSON.stringify({ ...p, _plik: undefined });
const czyNiezapisane = () => P && doPorownania(P) !== ZAPISANY;

function ustawProjekt(projekt) {
  P = projekt;
  ZAPISANY = doPorownania(P);
  api('ustaw_niezapisane', false);
  W = null;
}

function zmiana() {
  const nz = czyNiezapisane();
  api('ustaw_niezapisane', nz);
  rysujPasek();
}

/** Pyta o zapis niezapisanych zmian. true = można kontynuować (zapisano albo porzucono). */
async function pytajOZapis(pytanie = 'Kontynuować?') {
  const w = await modal({
    tytul: `${ik('uwaga')} Niezapisane zmiany`,
    tresc: `<p>Projekt <b>${esc(P.meta.project_no || nazwaPliku(P._plik) || 'bez nazwy')}</b> ma niezapisane zmiany. ${esc(pytanie)}</p>`,
    przyciski: [{ tekst: 'Anuluj', wartosc: null }, { tekst: 'Nie zapisuj', klasa: 'btn-czerwony', wartosc: 'porzuc' },
      { tekst: `${ik('zapisz')} Zapisz`, klasa: 'btn-hr', wartosc: 'zapisz', domyslny: true }],
  });
  if (w === 'porzuc') return true;
  if (w === 'zapisz') return zapiszProjekt(false);
  return false;
}

window.pytajPrzedZamknieciem = async () => {
  if (modalOtwarty) modalOtwarty.zamknij(null);
  if (await pytajOZapis('Zamknąć program?')) api('zamknij');
};

async function nowyProjekt() {
  if (czyNiezapisane() && !(await pytajOZapis('Utworzyć nowy projekt?'))) return;
  ustawProjekt(await api('nowy_projekt'));
  pokaz('obliczenia');
  toast('Nowy projekt.', 'ok', null, 2500);
}

async function otworzProjekt() {
  const ostatnie = (await api('ostatnie')) || [];
  const w = await modal({
    tytul: `${ik('otworz')} Otwórz projekt`,
    tresc: ostatnie.length ? `<p class="podp">Ostatnio otwierane</p><div class="lista-ostatnich">${ostatnie.map((o) => `
        <button data-wybor="${esc(o.sciezka)}"><b>${esc(o.nazwa)}</b><span>${esc(o.data)}</span><span style="grid-column:span 2">${esc(o.folder)}</span></button>`).join('')}</div>`
      : '<p class="podp">Brak ostatnio otwieranych projektów.</p>',
    przyciski: [{ tekst: 'Anuluj', wartosc: null }, { tekst: `${ik('folder')} Wybierz plik…`, klasa: 'btn-hr', wartosc: '', domyslny: true }],
  });
  if (w === null) return;
  if (czyNiezapisane() && !(await pytajOZapis('Otworzyć inny projekt?'))) return;
  const r = await api('wczytaj_projekt', w);
  if (r?.anulowano) return;
  if (!r?.ok) return toast(esc(r?.blad || 'Nie udało się wczytać projektu.'), 'blad');
  ustawProjekt(r.projekt);
  pokaz('obliczenia');
  toast(`Wczytano projekt <b>${esc(nazwaPliku(r.projekt._plik))}</b>`, 'ok', null, 3000);
  if (r.brakujace.length) toast(`Nie znaleziono plików załączników: ${r.brakujace.map(esc).join(', ')} - raport powstanie bez nich.`, 'blad');
}

async function zapiszProjekt(jako) {
  const r = await api('zapisz_projekt', P, jako);
  if (r?.anulowano) return false;
  if (!r?.ok) { toast(esc(r?.blad || 'Nie udało się zapisać projektu.'), 'blad'); return false; }
  P._plik = r.plik;
  ZAPISANY = doPorownania(P);
  api('ustaw_niezapisane', false);
  rysujPasek();
  toast(`Zapisano projekt <b>${esc(nazwaPliku(r.plik))}</b>`, 'ok', { tekst: 'Pokaż w folderze', fn: () => api('otworz_folder', r.plik) }, 4000);
  return true;
}

function rysujPasek() {
  const tytul = $('#pasek-tytul'), nad = $('#pasek-nad'), prawe = $('#pasek-prawe');
  odswiezDol();
  if (widok !== 'obliczenia') {
    nad.textContent = widok === 'ustawienia' ? 'Program' : 'Kalkulator Akustyki';
    tytul.textContent = { baza: 'Baza materiałów, silników i chłodnic', ustawienia: 'Ustawienia', zglos: 'Zgłoszenia' }[widok];
    prawe.innerHTML = '';
    return;
  }
  const nz = czyNiezapisane();
  nad.innerHTML = `${P._plik ? esc(nazwaPliku(P._plik)) : 'Nowy projekt'}${nz ? '<span class="stan-projektu"><span class="kropka-z"></span>niezapisane zmiany</span>' : P._plik ? '<span class="stan-projektu">zapisany</span>' : ''}`;
  tytul.innerHTML = `<span class="nazwa-projektu">${esc([P.meta.project_no, P.meta.client_name].filter(Boolean).join(' · ') || 'Analiza akustyczna obudowy')}</span>`;
  prawe.innerHTML = `
    <button class="btn btn-maly" id="p-nowy" title="Nowy projekt (Ctrl+N)">${ik('nowy')} Nowy</button>
    <button class="btn btn-maly" id="p-otworz" title="Otwórz projekt (Ctrl+O)">${ik('otworz')} Otwórz</button>
    <button class="btn btn-maly" id="p-zapisz-jako" title="Zapisz jako…">Zapisz jako…</button>
    <button class="btn btn-maly ${nz ? 'btn-hr' : ''}" id="p-zapisz" title="Zapisz projekt (Ctrl+S)">${ik('zapisz')} Zapisz</button>`;
  $('#p-nowy').onclick = nowyProjekt;
  $('#p-otworz').onclick = otworzProjekt;
  $('#p-zapisz').onclick = () => zapiszProjekt(false);
  $('#p-zapisz-jako').onclick = () => zapiszProjekt(true);
}

// ------------------------------------------------------------------ ekran obliczeń: pola projektu (ścieżki w obiekcie projektu, np. "meta.project_no")
const wez = (sc) => sc.split('.').reduce((o, k) => (o === undefined || o === null ? undefined : o[k]), P);
function ustaw(sc, v) {
  const k = sc.split('.');
  const ost = k.pop();
  k.reduce((o, x) => o[x], P)[ost] = v;
}
const MODULY = { sciana: 'Ściany', podloga: 'Podłoga / rama', sufit: 'Sufit / dach', drzwi: 'Drzwi', czerpnia: 'Czerpnia powietrza (wlot)', wyrzut: 'Wyrzut powietrza (wylot)', wydech: 'Wydech silnika (komin)' };
const AUTO_S = ['sciana', 'podloga', 'sufit'];
const Z_KIERUNKIEM = ['wydech', 'czerpnia', 'wyrzut'];
const TYPY_OBUDOWY = [['Obudowa kontenerowa', 'Obudowa kontenerowa'], ['Obudowa dźwiękochłonna', 'Obudowa dźwiękochłonna (canopy)'], ['Agregat otwarty', 'Agregat otwarty']];
const liczbaJs = (v) => { const n = Number(String(v ?? '').replace(',', '.').replace(/\s/g, '')); return Number.isFinite(n) ? n : 0; };
const sumaDb = (l) => { const w = l.reduce((s, x) => s + 10 ** (liczbaJs(x) / 10), 0); return w > 0 ? 10 * Math.log10(w) : 0; };
const widmoA = (l, j) => l.map((x, i) => liczbaJs(x) + (j === 'dB' ? INFO.korekta_a[i] : 0));

function pTekst(sc, etykieta, { podpowiedz = '', lista = '' } = {}) {
  return `<div class="p-pole"><label class="etykieta">${etykieta}</label><input class="pole" data-sc="${sc}" value="${esc(wez(sc))}" spellcheck="false" placeholder="${esc(podpowiedz)}" ${lista ? `list="${lista}"` : ''}></div>`;
}
function pLiczba(sc, etykieta, jedn = '', { podpowiedz = '' } = {}) {
  return `<div class="p-pole"><label class="etykieta">${etykieta}</label>
    <div class="z-jednostka"><input class="pole" data-sc="${sc}" value="${esc(wez(sc))}" inputmode="decimal" spellcheck="false" ${podpowiedz ? `placeholder="${esc(podpowiedz)}"` : ''}><span class="jedn">${jedn}</span></div></div>`;
}
function pSegment(sc, opcje, kl = '') {
  return `<div class="segment ${kl}" data-seg="${sc}">${opcje.map(([v, t]) => `<button data-v="${esc(v)}" class="${String(wez(sc)) === String(v) ? 'akt' : ''}">${t}</button>`).join('')}</div>`;
}
function pWybor(sc, opcje) {
  return `<select class="pole" data-sc="${sc}">${opcje.map(([v, t]) => `<option value="${esc(v)}" ${String(wez(sc)) === String(v) ? 'selected' : ''}>${esc(t)}</option>`).join('')}</select>`;
}
const SEG_JEDN = [['dBA', 'dB(A)'], ['dB', 'dB']];

/** Widmo w 8 pasmach: wiersz pól (ścieżka sc.0 … sc.7) i opcjonalnie wiersz wyliczony (np. LwA po korekcie A). */
function pWidmo(sc, etykieta, { jedn = null, wyliczone = null, opisWyl = 'LwA [dB(A)]' } = {}) {
  const w = wez(sc) || [];
  return `<div class="widmo" data-widmo="${sc}" data-jedn="${jedn || ''}">
    <div class="widmo-nag"><span class="etykieta">${etykieta}</span>${jedn ? `<span class="podp">jednostka</span>${pSegment(jedn, SEG_JEDN)}` : ''}</div>
    <div class="widmo-siatka"><span></span>${INFO.pasma.map((p) => `<span class="pasmo">${p}</span>`).join('')}<span class="pasmo">suma</span>
      <span class="et-wiersza">wpisane</span>${INFO.pasma.map((_, i) => `<input class="pole" data-sc="${sc}.${i}" value="${esc(w[i])}" inputmode="decimal">`).join('')}<span class="suma-w">${lp(sumaDb(w), 1)}</span>
      ${wyliczone ? `<span class="et-wiersza">${opisWyl}</span>${wyliczone.map((x) => `<span class="wyl">${lp(x, 1)}</span>`).join('')}<span class="suma-w"><b>${lp(sumaDb(wyliczone), 1)}</b></span>` : ''}
    </div></div>`;
}

// Zakładki ekranu obliczeń: dane projektu, obudowa, źródło, każdy moduł osobno, chłodnica. Wyniki po prawej - zawsze widoczne.
let zakladka = 'projekt';
const MODULY_KROTKO = { sciana: 'Ściany', podloga: 'Podłoga', sufit: 'Sufit', drzwi: 'Drzwi', czerpnia: 'Czerpnia', wyrzut: 'Wyrzut', wydech: 'Wydech' };
const zakladkiObl = () => [['projekt', 'Projekt', 'plik'], ['obudowa', 'Obudowa i punkty', 'obudowa'], ['silnik', 'Silnik', 'silnik'],
  ...INFO.moduly.map((k) => [k, MODULY_KROTKO[k], 'warstwy']), ['chlodnica', 'Chłodnica', 'wiatrak']];

/** Opis pod nazwą zakładki: moduł - udział w hałasie w pierwszym punkcie (albo "wył."), chłodnica - wł. / wył. */
function opisZakladki(z) {
  if (INFO.moduly.includes(z)) {
    if (!P.modules_active[z]) return 'wył.';
    const u = W?.ok ? W.punkty[0]?.moduly[z] : null;
    const k = z === 'wydech' && W?.ok ? W.punkty[0]?.komin : null;
    if (k) return `${lp(udzialKomin(k), 0)}% (komin)`;
    return u ? `${lp(u.procent, 0)}%` : '…';
  }
  if (z === 'chlodnica') return P.radiator.active ? 'wł.' : 'wył.';
  if (z === 'obudowa') return `${P.points.length} ${P.points.length === 1 ? 'punkt' : 'punkty'}`;
  if (z === 'silnik') return W?.ok ? `${lp(W.lw_a_suma, 1)} dB(A)` : '';
  return '';
}
function odswiezZakladki() {
  for (const [z] of zakladkiObl()) {
    const b = $(`[data-zakl="${z}"]`);
    if (!b) continue;
    b.querySelector('small').textContent = opisZakladki(z);
    const dom = W?.ok && W.punkty[0]?.dominujacy === z;
    b.classList.toggle('dominujacy', !!dom);
    b.classList.toggle('wylaczony', (INFO.moduly.includes(z) && !P.modules_active[z]) || (z === 'chlodnica' && !P.radiator.active));
  }
}

function rysujObliczenia() {
  const el = $('#widok-obliczenia');
  el.innerHTML = `<div class="dane" id="dane">
    <nav class="zakladki-obl">${zakladkiObl().map(([z, t, ikona]) => `<button data-zakl="${z}" class="${z === zakladka ? 'akt' : ''}">${ik(ikona)}<span>${t}</span><small></small></button>`).join('')}</nav>
    <div id="tresc-zakladki">${trescZakladki()}</div>
  </div>
  <aside class="wyniki"><div class="przewijane" id="wyniki"></div>
    <div class="stopka">
      <div class="przyciski-raportu"><button class="btn btn-hr btn-raport-duzy" id="generuj" title="Raport PDF (Ctrl+G)">${ik('dokument')} Raport PDF</button>
        <button class="btn btn-raport-duzy" id="generuj-docx" title="Raport do edycji w Wordzie (Ctrl+Shift+G)">${ik('word')} DOCX</button></div>
      <div class="podp" id="stopka-raportu"></div>
    </div>
  </aside>`;
  if (zakladka === 'projekt') rysujRewizje();
  odswiezZakladki();
  przelicz();
  $('#generuj').onclick = () => generujRaport('pdf');
  $('#generuj-docx').onclick = () => generujRaport('docx');
}

function trescZakladki() {
  const m = P.meta, ak = P.acoustics, rad = P.radiator;
  if (zakladka === 'projekt') {
    const silniki = Object.keys(BAZA.engines || {}).sort((a, b) => a.localeCompare(b, 'pl'));
    return `<div class="karta">
      <div class="naglowek-karty"><h3>${ik('plik')}Projekt</h3>${pSegment('meta.jezyk_raportu', [['Polski', 'Raport PL'], ['English', 'Raport EN']])}</div>
      <div class="pola">
        ${pTekst('meta.project_no', 'Numer projektu', { podpowiedz: 'np. PRJ-2026-224' })}${pTekst('meta.client_name', 'Klient')}${pTekst('meta.autor', 'Przygotował (inicjały)')}
        ${pTekst('meta.project_name', 'Tytuł raportu', { podpowiedz: 'Analiza akustyczna obudowy agregatu' })}${pTekst('meta.po_number', 'Numer P.O.')}${pTekst('meta.client_doc_no', 'Dokument klienta')}
        ${pTekst('meta.agregat_typ', 'Model agregatu', { podpowiedz: 'np. C3500D5EB' })}
        <div class="p-pole"><label class="etykieta"><span>Silnik</span>${m.silnik_typ && BAZA.engines?.[m.silnik_typ] ? '<span class="zbazy">z bazy</span>' : ''}</label>
          <input class="pole" id="f-silnik" data-sc="meta.silnik_typ" list="lista-silnikow" value="${esc(m.silnik_typ)}" spellcheck="false" placeholder="${silniki.length ? 'wybierz z bazy albo wpisz' : 'np. QSK78-G39'}">
          <datalist id="lista-silnikow">${silniki.map((n) => `<option value="${esc(n)}">${esc(BAZA.engines[n].genset || '')}</option>`).join('')}</datalist></div>
        <div class="p-pole"><label class="etykieta">Typ instalacji</label>${pWybor('meta.typ_obudowy', TYPY_OBUDOWY)}</div>
      </div>
      <p class="podp" style="margin:4px 0 0">Wybór silnika z bazy wpisuje jego widmo w zakładce <b>Silnik</b>.</p></div>
    <div class="karta"><div class="naglowek-karty"><h3>${ik('odswiez')}Tabela rewizji dokumentu</h3></div><div id="rewizje"></div></div>`;
  }
  if (zakladka === 'obudowa') {
    return `<div class="karta">
      <div class="naglowek-karty"><h3>${ik('obudowa')}Wymiary obudowy</h3><span class="podp">ściany, podłoga i sufit liczone z wymiarów</span></div>
      <div class="pola">${pLiczba('dimensions.kont_l', 'Długość', 'm')}${pLiczba('dimensions.kont_w', 'Szerokość', 'm')}${pLiczba('dimensions.kont_h', 'Wysokość', 'm')}</div></div>
    <div class="karta">
      <div class="naglowek-karty"><h3>${ik('info')}Punkty pomiarowe</h3><span class="podp">odległość od obudowy · pomiar 1,5 m nad terenem · ISO 3744</span></div>
      <table class="tabela-punktow"><thead><tr><th>Punkt pomiarowy</th><th>Odległość</th><th>Wymagany poziom</th><th></th></tr></thead><tbody>
        ${P.points.map((pt, i) => `<tr><td><input class="pole" data-sc="points.${i}.name" value="${esc(pt.name)}" placeholder="np. Front 1 m"></td>
          <td><div class="z-jednostka"><input class="pole" data-sc="points.${i}.d" value="${esc(pt.d)}" inputmode="decimal"><span class="jedn">m</span></div></td>
          <td><div class="z-jednostka"><input class="pole" data-sc="points.${i}.limit" value="${esc(pt.limit)}" inputmode="decimal"><span class="jedn">dB(A)</span></div></td>
          <td class="w-x">${P.points.length > 1 ? `<button class="btn btn-bez btn-ikona" data-usun-pkt="${i}" title="Usuń punkt">${ik('kosz')}</button>` : ''}</td></tr>`).join('')}
      </tbody></table>
      <button class="btn btn-maly" id="dodaj-pkt">${ik('plus')} Dodaj punkt</button></div>`;
  }
  if (zakladka === 'silnik') {
    return `<div class="karta">
      <div class="naglowek-karty"><h3>${ik('silnik')}Silnik - źródło hałasu</h3><span class="podp">${esc(m.silnik_typ || 'moc akustyczna Lw w pasmach oktawowych')}</span></div>
      ${pWidmo('acoustics.lw_vars', 'Widmo mocy akustycznej silnika (z karty producenta)', { jedn: 'acoustics.jedn_silnika', wyliczone: ak.jedn_silnika === 'dB' ? widmoA(ak.lw_vars, 'dB') : null, opisWyl: 'po korekcie A' })}
      <p class="podp" style="margin:6px 0 0">Karty silników podają zwykle widmo w <b>dB(A)</b> (A-weighted sound power) - wtedy program nie dodaje korekty A. Dla widma w dB korekta A jest dodawana w każdym paśmie.</p>
      <div class="zalaczniki-male" style="margin-top:12px">${zalacznikJeden('meta.engine_wiz_path', 'obraz', 'Obraz do raportu (widmo, karta)')}${zalacznikJeden('meta.engine_pdf_path', 'pdf', 'Karta silnika PDF (w rozdziale silnika w raporcie PDF)')}</div></div>
    <div class="karta">
      <div class="naglowek-karty"><h3>${ik('odswiez')}Pochłanianie wnętrza obudowy</h3><span class="podp">poziom ciśnienia wewnątrz: LpA = LwA + Room Abs.</span></div>
      ${pWidmo('acoustics.room_abs_vars', 'Room Abs. [dB]')}
      <div class="sabine">
        <span class="etykieta" style="grid-column:1/-1">Kalkulator pochłaniania (wzór Sabine'a)</span>
        ${pLiczba('sabine.alfa_podl', 'α podłogi', '')}${pLiczba('sabine.alfa_scian', 'α ścian i sufitu', '')}${pLiczba('sabine.sabine_q', 'Kierunkowość Q', '')}${pLiczba('sabine.sabine_r', 'Odległość r', 'm')}
        <button class="btn btn-maly" id="sabine-oblicz">${ik('odswiez')} Oblicz i wpisz</button>
      </div></div>`;
  }
  if (zakladka === 'chlodnica') {
    return `<div class="karta">
      <div class="naglowek-karty"><h3>${ik('wiatrak')}Chłodnica na wyrzucie</h3>
        <label class="zaznacz"><input type="checkbox" data-sc-bool="radiator.active" ${rad.active ? 'checked' : ''}> chłodnica aktywna</label></div>
      ${rad.active ? `<div class="pola p2">
          <div class="p-pole"><label class="etykieta">Z bazy</label><select class="pole" id="f-chlodnica"><option value="">— wybierz, aby wpisać dane —</option>${Object.keys(BAZA.radiators || {}).sort().map((n) => `<option ${n === rad.name ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select></div>
          ${pTekst('radiator.name', 'Nazwa w raporcie')}</div>
        ${pWidmo('radiator.lw', 'Hałas własny chłodnicy Lw', { jedn: 'radiator.jedn', wyliczone: rad.jedn === 'dB' ? widmoA(rad.lw, 'dB') : null, opisWyl: 'po korekcie A' })}
        ${pWidmo('radiator.att', 'Tłumienie rdzenia chłodnicy (hałas silnika przez rdzeń) [dB]')}
        ${!P.modules_active.wyrzut ? `<div class="uwaga-pola">${ik('uwaga')} Chłodnica jest liczona na module „Wyrzut” - włącz go w zakładce Wyrzut.</div>` : ''}
        <div class="zalaczniki-male" style="margin-top:12px">${zalacznikJeden('radiator.wiz_path', 'obraz', 'Obraz do raportu')}${zalacznikJeden('radiator.pdf_path', 'pdf', 'Karta chłodnicy PDF')}</div>`
      : '<p class="podp" style="margin:0">Włącz, gdy wyrzut powietrza przechodzi przez chłodnicę - program doda hałas własny chłodnicy i hałas silnika przechodzący przez rdzeń.</p>'}</div>`;
  }
  return kartaModulu(zakladka);
}

function kartaModulu(k) {
  const akt = !!P.modules_active[k], md = P.mod_dims[k], warstwy = P.layers[k] || [], zal = P.mod_attachments[k] || [];
  const s = W?.ok ? W.powierzchnie[k] : null;
  const u = W?.ok ? W.punkty[0]?.moduly[k] : null;
  return `<div class="karta modul-karta" data-modul="${k}">
    <div class="naglowek-karty"><h3>${ik('warstwy')}${MODULY[k]}</h3>
      <label class="zaznacz"><input type="checkbox" data-sc-bool="modules_active.${k}" ${akt ? 'checked' : ''}> uwzględnij w obliczeniach</label></div>
    ${!akt ? `<p class="podp" style="margin:0">Moduł wyłączony - nie jest liczony ani opisany w raporcie.${k === 'wydech' ? ' Włącz, gdy komin wydechu wychodzi z obudowy.' : ''}</p></div>` : `
    <div class="podsumowanie-modulu"><span>Powierzchnia <b class="s-modulu">${k === 'wydech' ? '—' : s === null ? '…' : `${lp(s)} m²`}</b></span>
      <span>Izolacyjność śr. <b class="r-modulu">${W?.ok && W.moduly[k] ? `${lp(W.moduly[k].r.reduce((a, b) => a + b, 0) / 8, 1)} dB` : '…'}</b></span>
      <span>Udział w punkcie ${esc(P.points[0]?.name || '')} <b class="u-modulu">${u ? opisUdzialu(k) : '…'}</b></span></div>
    <div class="tresc-modulu">
      ${AUTO_S.includes(k) ? `<p class="podp" style="margin:0">Powierzchnia z wymiarów obudowy (zakładka Obudowa)${k === 'sciana' ? ' minus drzwi, czerpnia i wyrzut' : ''}.</p>`
        : `<div class="pola p4">${k !== 'wydech' ? `${pLiczba(`mod_dims.${k}.w`, 'Szerokość', 'm')}${pLiczba(`mod_dims.${k}.h`, 'Wysokość', 'm')}` : ''}${pLiczba(`mod_dims.${k}.qty`, k === 'wydech' ? 'Liczba kominów' : 'Ilość', 'szt.')}</div>`}
      ${k === 'wydech' ? polaWydechu() : Z_KIERUNKIEM.includes(k) ? `<div class="pola p4"><div class="p-pole"><label class="etykieta">Kierunek wylotu</label>${pSegment(`mod_dims.${k}.kierunek`, [['Poziomy', 'Poziomo'], ['Pionowy w górę', 'W górę']])}</div>
        ${pLiczba(`mod_dims.${k}.wysokosc`, 'Wysokość nad punktem', 'm', { podpowiedz: '0 = na wysokości punktu' })}</div>` : ''}
    </div></div>
    <div class="karta">
      <div class="naglowek-karty"><h3>${ik('warstwy')}Warstwy ${k === 'wydech' ? '(tłumiki spalin)' : 'izolacyjne'}</h3><span class="podp">z bazy materiałów albo własne - sumowane energetycznie</span></div>
      <datalist id="lista-materialow">${Object.keys(BAZA.materials || {}).sort((a, b) => a.localeCompare(b, 'pl')).map((n) => `<option value="${esc(n)}">`).join('')}</datalist>
      <div class="warstwy">${warstwy.map((w, i) => wierszWarstwy(k, w, i, warstwy.length)).join('') || `<p class="podp" style="margin:4px 0">${k === 'wydech' ? 'Brak tłumika - dodaj tłumik spalin jako warstwę.' : 'Brak warstw - moduł liczony jak otwór (bez izolacyjności).'}</p>`}</div>
      <div class="akcje-modulu" style="margin-top:10px"><button class="btn btn-maly" data-dodaj-warstwe="${k}">${ik('plus')} Dodaj warstwę</button></div></div>
    <div class="karta">
      <div class="naglowek-karty"><h3>${ik('obraz')}Załączniki do raportu</h3><span class="podp">obrazy w opisie modułu, karty PDF w raporcie PDF zaraz po rozdziale</span></div>
      <div class="akcje-modulu"><button class="btn btn-maly" data-zal-modulu="${k}" data-typ="obraz">${ik('obraz')} Dodaj obraz</button>
        <button class="btn btn-maly" data-zal-modulu="${k}" data-typ="pdf">${ik('dokument')} Dodaj kartę PDF</button></div>
      ${zal.length ? `<div class="lista-zalacznikow" style="margin-top:10px">${zal.map((a, i) => `<span class="znacznik akt ${MINIATURY_BRAK.has(a.path) ? 'brak' : ''}">${ik(a.type === 'Dokument (PDF)' ? 'dokument' : 'obraz')} ${esc(nazwaPliku(a.path))}<button class="btn btn-bez btn-ikona" data-usun-zal-modulu="${k}" data-i="${i}" title="Usuń">${ik('x')}</button></span>`).join('')}</div>` : ''}`}
  </div>`;
}
const MINIATURY_BRAK = new Set();

// Wydech: położenie wylotu (wynik po stronie komina - źródło punktowe) i widmo jako Lw albo Lp@1m z karty tłumika (+11 dB)
const LP_NA_LW = 10 * Math.log10(4 * Math.PI);
function widmoWydechuA() {
  const ak = P.acoustics, dodaj = ak.wielkosc_wydechu === 'Lp1m' ? LP_NA_LW : 0;
  return widmoA(ak.exhaust_lw_vars.map((x) => liczbaJs(x) + dodaj), ak.jedn_wydechu);
}
function polaWydechu() {
  const ak = P.acoustics, lp1m = ak.wielkosc_wydechu === 'Lp1m', pusty = !ak.exhaust_lw_vars.some((x) => liczbaJs(x));
  const stary = String(P.mod_dims.wydech.nad_dachem ?? '').trim() === '';
  return `<div class="pola p4"><div class="p-pole"><label class="etykieta">Kierunek wylotu</label>${pSegment('mod_dims.wydech.kierunek', [['Poziomy', 'Poziomo'], ['Pionowy w górę', 'W górę']])}</div>
      ${pLiczba('mod_dims.wydech.nad_dachem', 'Wylot nad dachem obudowy', 'm', { podpowiedz: 'np. 1,8' })}
      ${pLiczba('mod_dims.wydech.od_sciany', 'Wylot od najbliższej ściany', 'm', { podpowiedz: 'poziomo, np. 1,2' })}
    </div>
    <p class="podp" style="margin:2px 0 8px">Wylot = miejsce, gdzie spaliny wychodzą w powietrze (koniec kolana / komina). Punkt pomiaru: 1,5 m nad terenem. Program liczy wydech dodatkowo jako źródło punktowe
      w punkcie po stronie komina - o zgodności decyduje gorszy z dwóch wyników.<span class="geo-wydechu">${opisGeoWydechu()}</span></p>
    ${stary ? `<div class="uwaga-pola">${ik('info')} Projekt sprzed wersji 2.0.2 - podaj położenie wylotu, żeby policzyć poziom po stronie komina (teraz tylko średnia wokół obudowy, wysokość nad punktem ${lc(liczbaJs(P.mod_dims.wydech.wysokosc))} m).</div>` : ''}
    <div class="widmo-nag" style="margin-top:6px"><span class="etykieta">Widmo hałasu wydechu przed tłumikiem</span>${pSegment('acoustics.wielkosc_wydechu', [['Lp1m', 'Lp@1m (karta tłumika)'], ['Lw', 'Lw (moc)']])}</div>
    ${pWidmo('acoustics.exhaust_lw_vars', lp1m ? 'Lp w odległości 1 m od wylotu - bez tłumika (karta producenta tłumika)' : 'Moc akustyczna wydechu Lw - bez tłumika',
      { jedn: 'acoustics.jedn_wydechu', wyliczone: lp1m || ak.jedn_wydechu === 'dB' ? widmoWydechuA() : null, opisWyl: lp1m ? 'LwA = Lp + 11' : 'po korekcie A' })}
    <p class="podp" style="margin:4px 0 0">Pasma 63 Hz - 8 kHz. Gdy karta zaczyna się od 31,5 Hz - pomiń tę kolumnę. Kolumna „dB” → jednostka dB, kolumna „dB(A)” → dB(A).${lp1m ? ' Lp@1m przeliczane na moc: Lw = Lp + 10·lg(4π) ≈ Lp + 11 dB.' : ''}</p>
    <div class="uwaga-pola uwaga-pusty-wydech" ${pusty ? '' : 'hidden'}>${ik('uwaga')} Widmo wydechu jest puste (same zera) - wydech nie jest liczony. Wpisz widmo z karty silnika albo tłumika.</div>`;
}
const udzialKomin = (k) => (k.wynik ? 10 ** ((k.wydech - k.wynik) / 10) * 100 : 0);
/** Udział modułu w pierwszym punkcie; wydech z położeniem wylotu - po stronie komina i średnio. */
function opisUdzialu(k) {
  const pt = W?.ok ? W.punkty[0] : null, u = pt?.moduly[k];
  if (!u) return '—';
  if (k === 'wydech' && pt.komin) return `${lp(pt.komin.wydech, 1)} dB(A) · ${lp(udzialKomin(pt.komin), 0)}% przy kominie (średnio ${lp(u.udzial, 1)})`;
  return `${lp(u.udzial, 1)} dB(A) · ${lp(u.procent, 0)}%`;
}
function opisGeoWydechu() {
  const k = W?.ok ? W.punkty[0]?.komin : null;
  const m1 = W?.ok && W.wydech_lpa_1m_suma ? ` Wydech 1 m od wylotu (za tłumikiem): <b>${lp(W.wydech_lpa_1m_suma, 1)} dB(A)</b>.` : '';
  return m1 + (k ? ` Punkt ${esc(P.points[0]?.name || '')}: wylot ${lp(k.wysokosc, 1)} m nad punktem, ${lp(k.r, 1)} m od punktu → wydech <b>${lp(k.wydech, 1)} dB(A)</b>, łącznie po stronie komina <b>${lp(k.wynik, 1)} dB(A)</b>.` : '');
}

function wierszWarstwy(k, w, i, ile) {
  const sc = `layers.${k}.${i}`;
  const zBazy = BAZA.materials?.[w.name];
  return `<div class="warstwa">
    <div class="warstwa-gora">
      <input class="pole nazwa-warstwy" data-sc="${sc}.name" data-warstwa="${k}.${i}" list="lista-materialow" value="${esc(w.name)}" placeholder="materiał z bazy albo własna nazwa">
      ${zBazy ? '<span class="zbazy">z bazy</span>' : ''}
      <select class="pole typ-warstwy" data-sc="${sc}.type">${INFO.typy_materialu.map((t) => `<option ${w.type === t ? 'selected' : ''}>${t}</option>`).join('')}</select>
      <div class="z-jednostka grubosc"><input class="pole" data-sc="${sc}.d" value="${esc(w.d)}" inputmode="decimal" title="Grubość / długość"><span class="jedn">mm</span></div>
      <span class="akcje">${i ? `<button class="btn btn-bez btn-ikona" data-warstwa-gora="${k}.${i}" title="W górę">▲</button>` : ''}${i < ile - 1 ? `<button class="btn btn-bez btn-ikona" data-warstwa-dol="${k}.${i}" title="W dół">▼</button>` : ''}
        <button class="btn btn-bez btn-ikona" data-usun-warstwe="${k}.${i}" title="Usuń warstwę">${ik('kosz')}</button></span>
    </div>
    ${w.type === 'Rw / Ctr' ? `<div class="pola p4">${pLiczba(`${sc}.rw`, 'Rw', 'dB')}${pLiczba(`${sc}.ctr`, 'Ctr', 'dB')}<p class="podp" style="grid-column:span 2; margin:auto 0">Rw + Ctr przyjmowane we wszystkich pasmach (zachowawczo w 63-125 Hz zawyża izolacyjność).</p></div>`
      : `<div class="widmo-siatka mala">${INFO.pasma.map((p) => `<span class="pasmo">${p}</span>`).join('')}${INFO.pasma.map((_, j) => `<input class="pole" data-sc="${sc}.spec.${j}" value="${esc(w.spec[j])}" inputmode="decimal">`).join('')}</div>`}
  </div>`;
}

function zalacznikJeden(sc, typ, etykieta) {
  const p = wez(sc);
  return `<div class="zal-jeden ${p ? 'jest' : ''}">${ik(typ === 'pdf' ? 'dokument' : 'obraz')}<div><b>${etykieta}</b><span class="nazwa-pliku" title="${esc(p)}">${p ? esc(nazwaPliku(p)) : 'brak'}</span></div>
    <button class="btn btn-maly" data-zal-wybierz="${sc}" data-typ="${typ}">${p ? 'Zmień' : 'Wybierz…'}</button>${p ? `<button class="btn btn-maly btn-ikona" data-zal-usun="${sc}" title="Usuń">${ik('x')}</button>` : ''}</div>`;
}

function przerysujDane() {
  const t = $('#tresc-zakladki');
  if (!t) return rysujObliczenia();
  const przew = $('#dane').scrollTop;
  t.innerHTML = trescZakladki();
  if (zakladka === 'projekt') rysujRewizje();
  $('#dane').scrollTop = przew;
  odswiezZakladki();
  przeliczPozniej(0);
}
function pokazZakladke(z) {
  zakladka = z;
  $$('.zakladki-obl button').forEach((b) => b.classList.toggle('akt', b.dataset.zakl === z));
  przerysujDane();
  $('#dane').scrollTop = 0;
}

const widokObl = $('#widok-obliczenia');
widokObl.addEventListener('input', (e) => {
  const sc = e.target.dataset.sc;
  if (sc && e.target.type !== 'checkbox' && e.target.tagName !== 'SELECT') {
    ustaw(sc, e.target.value);
    if (['meta.project_no', 'meta.client_name'].includes(sc)) rysujPasek();
    zmiana();
    przeliczPozniej();
  }
  const rev = e.target.dataset.rev;
  if (rev !== undefined) { P.revisions[+rev][e.target.dataset.k] = e.target.value; zmiana(); }
});
widokObl.addEventListener('change', (e) => {
  const t = e.target;
  if (t.dataset.scBool) {
    ustaw(t.dataset.scBool, t.checked);
    zmiana(); przerysujDane();
    return;
  }
  if (t.tagName === 'SELECT' && t.dataset.sc) {
    ustaw(t.dataset.sc, t.value); zmiana();
    if (t.dataset.sc.endsWith('.type')) przerysujDane(); else przeliczPozniej(0);
    return;
  }
  if (t.id === 'f-silnik' && BAZA.engines?.[t.value]) { zastosujSilnik(t.value); return; }
  if (t.id === 'f-chlodnica' && t.value) { zastosujChlodnice(t.value); return; }
  if (t.dataset.warstwa) {
    const [k, i] = t.dataset.warstwa.split('.');
    if (BAZA.materials?.[t.value]) { P.layers[k][+i] = warstwaZBazy(t.value); zmiana(); przerysujDane(); }
  }
});
widokObl.addEventListener('click', async (e) => {
  const zb = e.target.closest('[data-zakl]');
  if (zb) { pokazZakladke(zb.dataset.zakl); return; }
  const b = e.target.closest('[data-seg] button');
  if (b) {
    ustaw(b.closest('[data-seg]').dataset.seg, b.dataset.v);
    zmiana(); przerysujDane();
    return;
  }
  const t = e.target.closest('button');
  if (!t) return;
  const ds = t.dataset;
  if (t.id === 'dodaj-pkt') { P.points.push({ name: '', d: 1, limit: P.points.at(-1)?.limit ?? 85 }); zmiana(); przerysujDane(); }
  else if (ds.usunPkt !== undefined) { P.points.splice(+ds.usunPkt, 1); zmiana(); przerysujDane(); }
  else if (t.id === 'sabine-oblicz') {
    const d = P.dimensions, s = P.sabine;
    const r = await api('sabine', d.kont_l, d.kont_w, d.kont_h, s.alfa_podl, s.alfa_scian, s.sabine_q, s.sabine_r);
    if (!r?.ok) return toast(esc(r?.blad || 'Błąd obliczeń pochłaniania.'), 'blad');
    P.acoustics.room_abs_vars = Array(8).fill(Math.round(r.room_abs * 100) / 100);
    zmiana(); przerysujDane();
    toast(`Pochłanianie wnętrza: <b>${lp(r.room_abs)} dB</b> (ᾱ = ${lp(r.alfa, 3)}, R = ${lp(r.stala, 1)} m²) - wpisane we wszystkich pasmach.`, 'ok', null, 6000);
  } else if (ds.dodajWarstwe) { P.layers[ds.dodajWarstwe].push(nowaWarstwa()); zmiana(); przerysujDane(); $$(`[data-modul="${ds.dodajWarstwe}"] .nazwa-warstwy`).at(-1)?.focus(); }
  else if (ds.usunWarstwe) { const [k, i] = ds.usunWarstwe.split('.'); P.layers[k].splice(+i, 1); zmiana(); przerysujDane(); }
  else if (ds.warstwaGora || ds.warstwaDol) {
    const [k, i] = (ds.warstwaGora || ds.warstwaDol).split('.'), a = +i, b2 = ds.warstwaGora ? a - 1 : a + 1, l = P.layers[k];
    [l[a], l[b2]] = [l[b2], l[a]]; zmiana(); przerysujDane();
  } else if (ds.zalWybierz) {
    const s = await api('wybierz_plik', ds.typ); if (s) { ustaw(ds.zalWybierz, s); zmiana(); przerysujDane(); }
  } else if (ds.zalUsun) { ustaw(ds.zalUsun, ''); zmiana(); przerysujDane(); }
  else if (ds.zalModulu) {
    const s = await api('wybierz_plik', ds.typ);
    if (s) { P.mod_attachments[ds.zalModulu].push({ type: ds.typ === 'pdf' ? 'Dokument (PDF)' : 'Obraz (PNG/JPG)', path: s }); zmiana(); przerysujDane(); }
  } else if (ds.usunZalModulu) { P.mod_attachments[ds.usunZalModulu].splice(+ds.i, 1); zmiana(); przerysujDane(); }
});

const nowaWarstwa = () => ({ name: '', name_en: '', desc_pl: '', desc_en: '', type: 'Widmo pasmowe', d: 0, rw: 0, ctr: 0, spec: Array(8).fill(0) });
function warstwaZBazy(nazwa) {
  const u = BAZA.materials[nazwa];
  return { ...nowaWarstwa(), name: nazwa, name_en: u.name_en || '', desc_pl: u.desc_pl || '', desc_en: u.desc_en || '', type: u.type || 'Widmo pasmowe',
    d: u.d ?? 0, rw: u.rw ?? 0, ctr: u.ctr ?? 0, spec: [...(u.spec || []), 0, 0, 0, 0, 0, 0, 0, 0].slice(0, 8) };
}

function zastosujSilnik(nazwa) {
  const u = BAZA.engines[nazwa];
  P.meta.silnik_typ = nazwa;
  if (u.genset && !P.meta.agregat_typ) P.meta.agregat_typ = u.genset;
  P.acoustics.lw_vars = [...(u.lw || []), 0, 0, 0, 0, 0, 0, 0, 0].slice(0, 8);
  P.acoustics.jedn_silnika = u.jedn || 'dBA';
  plikiZBazy(u, 'meta.engine_wiz_path', 'meta.engine_pdf_path');
  zmiana(); przerysujDane();
  toast(`Silnik z bazy: <b>${esc(nazwa)}</b> - widmo ${u.jedn === 'dB' ? 'w dB (z korektą A)' : 'w dB(A)'} wpisane.`, 'ok', null, 3500);
}
function zastosujChlodnice(nazwa) {
  const u = BAZA.radiators[nazwa];
  Object.assign(P.radiator, { active: true, name: nazwa, lw: [...(u.lw || []), 0, 0, 0, 0, 0, 0, 0, 0].slice(0, 8), att: [...(u.att || []), 0, 0, 0, 0, 0, 0, 0, 0].slice(0, 8), jedn: u.jedn || 'dB' });
  plikiZBazy(u, 'radiator.wiz_path', 'radiator.pdf_path');
  zmiana(); przerysujDane();
}

let ostrzezenieOLogowaniu = false;
/** Karta / obraz producenta z bazy jako załącznik projektu (pliki producentów - po zalogowaniu). */
function plikiZBazy(u, scObraz, scPdf) {
  for (const [rola, sc] of [['obraz', scObraz], ['karta', scPdf]]) {
    const p = u.pliki?.[rola];
    if (!p) continue;
    api('plik_bazy', p.id).then((r) => {
      if (r?.ok) { ustaw(sc, r.sciezka); zmiana(); if (widok === 'obliczenia') przerysujDane(); }
      else if (r?.zaloguj && !ostrzezenieOLogowaniu) { ostrzezenieOLogowaniu = true; toast('Karty producentów z bazy są dostępne po zalogowaniu.', '', { tekst: 'Zaloguj', fn: () => ekranLogowania() }, 7000); }
    });
  }
}

function rysujRewizje() {
  const el = $('#rewizje');
  if (!el) return;
  el.innerHTML = `<table class="rewizje"><thead><tr><th>Rew.</th><th>Opis zmiany</th><th>Wykonał</th><th>Sprawdził</th><th>Data</th><th></th></tr></thead><tbody>
    ${P.revisions.map((r, i) => `<tr>
      <td class="w-rev"><input class="pole" data-rev="${i}" data-k="rev" value="${esc(r.rev)}"></td>
      <td><input class="pole" data-rev="${i}" data-k="desc" value="${esc(r.desc)}"></td>
      <td class="w-kr"><input class="pole" data-rev="${i}" data-k="drawn" value="${esc(r.drawn)}"></td>
      <td class="w-kr"><input class="pole" data-rev="${i}" data-k="check" value="${esc(r.check)}"></td>
      <td class="w-data"><input class="pole" data-rev="${i}" data-k="date" value="${esc(r.date)}"></td>
      <td class="w-x"><button class="btn btn-bez btn-ikona" data-usun-rev="${i}" title="Usuń rewizję">${ik('kosz')}</button></td></tr>`).join('')}
    </tbody></table>
    <button class="btn btn-maly" id="dodaj-rev">${ik('plus')} Dodaj rewizję</button>`;
  $('#dodaj-rev').onclick = () => {
    const nr = P.revisions.length ? String((parseInt(P.revisions.at(-1).rev, 10) || 0) + 1) : '0';
    P.revisions.push({ rev: nr, desc: '', drawn: P.meta.autor, check: '', date: new Date().toISOString().slice(0, 10) });
    zmiana(); rysujRewizje(); odswiezLicznikRewizji();
    $$('#rewizje [data-k="desc"]').at(-1)?.focus();
  };
  $$('[data-usun-rev]', el).forEach((b) => b.onclick = () => { P.revisions.splice(+b.dataset.usunRev, 1); zmiana(); rysujRewizje(); odswiezLicznikRewizji(); });
}
function odswiezLicznikRewizji() { const s = $('#sekcja-rewizje summary .podp'); if (s) s.textContent = `(${P.revisions.length})`; }

// ------------------------------------------------------------------ obliczenia i wyniki
function przeliczPozniej(ms = 160) {
  clearTimeout(przeliczTimer);
  przeliczTimer = setTimeout(przelicz, ms);
}

async function przelicz() {
  W = await api('oblicz', P);
  if (!W) return;
  rysujWyniki();
  odswiezZakladki();
  const k = zakladka;
  if (W.ok && INFO.moduly.includes(k) && P.modules_active[k]) {
    const u = W.punkty[0]?.moduly[k], r = W.moduly[k]?.r;
    const ust = (sel, t) => { const x = $(sel); if (x) x.textContent = t; };
    ust('.s-modulu', k === 'wydech' ? '—' : `${lp(W.powierzchnie[k])} m²`);
    ust('.r-modulu', r ? `${lp(r.reduce((a, b) => a + b, 0) / 8, 1)} dB` : '—');
    ust('.u-modulu', u ? opisUdzialu(k) : '—');
    const g = $('.geo-wydechu');
    if (g) g.innerHTML = opisGeoWydechu();
    const pw = $('.uwaga-pusty-wydech');
    if (pw) pw.hidden = !W.wydech_pusty;
  }
  // wiersze wyliczone pod widmami (LwA po korekcie A / Lp + 11) - na bieżąco przy wpisywaniu
  $$('.widmo[data-widmo]').forEach((el) => {
    const wyl = el.querySelectorAll('.wyl');
    if (!wyl.length) return;
    const sc = el.dataset.widmo;
    const v = sc === 'acoustics.exhaust_lw_vars' ? widmoWydechuA() : widmoA(wez(sc) || [], el.dataset.jedn ? wez(el.dataset.jedn) : 'dBA');
    wyl.forEach((s, i) => { s.textContent = lp(v[i], 1); });
    const sumy = el.querySelectorAll('.suma-w');
    if (sumy[0]) sumy[0].textContent = lp(sumaDb(wez(sc) || []), 1);
    if (sumy[1]) sumy[1].innerHTML = `<b>${lp(sumaDb(v), 1)}</b>`;
  });
}

function rysujWyniki() {
  const el = $('#wyniki');
  if (!el) return;
  const stopka = $('#stopka-raportu');
  if (!W?.ok) {
    el.innerHTML = `<div class="werdykt brak"><div class="ikona-w">${ik('info')}</div><div><b>Uzupełnij dane</b><span>Wyniki pojawią się od razu po wpisaniu brakujących wartości.</span></div></div>
      <div class="bledy-danych">Brakuje albo niepoprawne:<ul>${Object.values(W?.bledy || {}).map((t) => `<li>${esc(t)}</li>`).join('')}</ul></div>`;
    $('#generuj').disabled = $('#generuj-docx').disabled = true;
    stopka.textContent = 'Raport można wygenerować po uzupełnieniu danych.';
    return;
  }
  $('#generuj').disabled = $('#generuj-docx').disabled = false;
  const zal = [P.meta.engine_wiz_path, P.meta.engine_pdf_path, ...(W.chlodnica ? [P.radiator.wiz_path, P.radiator.pdf_path] : []),
    ...Object.keys(W.moduly).flatMap((k) => (P.mod_attachments[k] || []).map((a) => a.path))].filter(Boolean).length;
  stopka.textContent = `Raport ${P.meta.jezyk_raportu === 'English' ? 'po angielsku' : 'po polsku'} · załączniki: ${zal} · ${opisSzablonu()}`;
  const zle = W.punkty.filter((p) => !p.zgodny);
  el.innerHTML = `
    <div class="werdykt ${zle.length ? 'zle' : 'ok'}"><div class="ikona-w">${ik(zle.length ? 'uwaga' : 'ok')}</div>
      <div><b>${zle.length ? 'Przekroczony poziom hałasu' : 'Obudowa spełnia wymagania'}</b><span>${zle.length ? `w punktach: ${zle.map((p) => esc(p.name || `${lc(p.d)} m`)).join(', ')}` : `${W.punkty.length} ${W.punkty.length === 1 ? 'punkt pomiarowy' : 'punkty pomiarowe'} · LwA silnika ${lp(W.lw_a_suma, 1)} dB(A)`}</span></div></div>
    ${W.punkty.map((pt) => `<div class="wynik">
      <div class="gora-wyniku"><div><div class="tytul-wyniku">${ik('info')}${esc(pt.name || `${lc(pt.d)} m`)} <span class="podp">· ${lc(pt.d)} m</span></div>
        <div class="wartosc">${lp(pt.decyduje, 1)}<small>dB(A)</small></div>
        ${pt.komin ? `<div class="wartosc-opis">${pt.komin.wynik >= pt.wynik ? '<b>po stronie komina</b>' : 'po stronie komina'} ${lp(pt.komin.wynik, 1)} · ${pt.komin.wynik < pt.wynik ? '<b>średnio wokół obudowy</b>' : 'średnio wokół obudowy'} ${lp(pt.wynik, 1)}</div>` : ''}</div>
        ${pt.zgodny ? `<span class="status ok">${ik('ok')}≤ ${lc(pt.limit)} · zapas ${lp(pt.zapas, 1)} dB</span>` : `<span class="status zle">${ik('uwaga')}+${lp(-pt.zapas, 1)} dB ponad ${lc(pt.limit)}</span>`}</div>
      <div class="udzialy">${Object.entries(pt.moduly).sort((a, b) => b[1].udzial - a[1].udzial).map(([k, m], i) => `
        <div class="udzial ${i === 0 ? 'glowny' : ''}"><span>${MODULY[k]}</span><div class="pasek-u"><i style="width:${Math.max(2, m.procent)}%"></i></div><b>${lp(m.udzial, 1)}</b><span class="podp">${lp(m.procent, 0)}%</span></div>`).join('')}</div>
    </div>`).join('')}
    <div class="wynik"><div class="tytul-wyniku">${ik('silnik')}Źródło i wnętrze obudowy</div>
      ${wykresWidma()}
      <div class="szczegol">LwA silnika <b>${lp(W.lw_a_suma, 1)} dB(A)</b> · LpA wewnątrz <b>${lp(W.lp_a_suma, 1)} dB(A)</b> · powierzchnia modułów ${lp(W.s_calkowita, 1)} m²${W.chlodnica ? ` · chłodnica LwA ${lp(W.rad_lw_a_suma, 1)} dB(A)` : ''}${W.wydech_lw_a_suma ? ` · wydech LwA ${lp(W.wydech_lw_a_suma, 1)} dB(A)` : ''}</div></div>`;
}

/** Widmo w pasmach: LwA silnika, LpA wewnątrz i poziom w pierwszym punkcie. */
function wykresWidma() {
  const pt = W.punkty[0], szer = 330, wys = 150, l = 30, pr = 8, g = 10, d = 22;
  const serie = [[W.lw_a, 'var(--kolor)', ''], [W.lp_a, 'var(--muted)', '4 3'], [pt.pasma, 'var(--zielony)', '']];
  const w = serie.flatMap((s) => s[0]).filter((x) => x);
  const dol = Math.floor(Math.min(...w) / 10) * 10, gora = Math.max(dol + 10, Math.ceil(Math.max(...w) / 10) * 10);
  const x = (i) => l + (i + 0.5) / 8 * (szer - l - pr), y = (v) => g + (1 - (v - dol) / (gora - dol)) * (wys - g - d);
  const siatka = []; for (let v = dol; v <= gora; v += 10) siatka.push(v);
  return `<div class="wykres"><svg viewBox="0 0 ${szer} ${wys}" role="img" aria-label="Widmo hałasu">
    ${siatka.map((v) => `<line class="siatka-w" x1="${l}" x2="${szer - pr}" y1="${y(v)}" y2="${y(v)}"/><text class="opis-w" x="${l - 5}" y="${y(v) + 3}" text-anchor="end">${v}</text>`).join('')}
    ${INFO.pasma.map((p, i) => `<text class="opis-w" x="${x(i)}" y="${wys - 7}" text-anchor="middle">${p}</text>`).join('')}
    <line class="os-w" x1="${l}" x2="${szer - pr}" y1="${wys - d}" y2="${wys - d}"/>
    ${serie.map(([s, k, kr]) => `<path d="${s.map((v, i) => `${i ? 'L' : 'M'}${x(i).toFixed(1)},${y(v).toFixed(1)}`).join('')}" fill="none" stroke="${k}" stroke-width="2.2" ${kr ? `stroke-dasharray="${kr}"` : ''}/>${s.map((v, i) => `<circle cx="${x(i).toFixed(1)}" cy="${y(v).toFixed(1)}" r="2.6" fill="${k}"/>`).join('')}`).join('')}
  </svg>
  <div class="legenda"><span><i style="background:var(--kolor)"></i>LwA silnika</span><span><i style="background:var(--muted)"></i>LpA wewnątrz</span><span><i style="background:var(--zielony)"></i>${esc(pt.name || `${lc(pt.d)} m`)}</span></div></div>`;
}

async function generujRaport(format = 'pdf') {
  if (!W?.ok) return toast('Uzupełnij dane przed wygenerowaniem raportu.', 'blad');
  const b = $(format === 'pdf' ? '#generuj' : '#generuj-docx'), tekst = b.innerHTML;
  b.disabled = true;
  if (format === 'pdf') b.innerHTML = 'Tworzenie PDF…';
  const r = await api('generuj_raport', P, format);
  b.disabled = false;
  b.innerHTML = tekst;
  if (r?.anulowano) return;
  if (!r?.ok) return toast(esc(r?.blad || 'Nie udało się utworzyć raportu.'), 'blad');
  toast(`Zapisano raport <b>${esc(nazwaPliku(r.plik))}</b>${r.karty ? `<br><span class="podp">Dołączone karty PDF: ${r.karty}</span>` : ''}${r.pominiete.length ? `<br><span class="podp">Pominięto: ${r.pominiete.map(esc).join(', ')}</span>` : ''}`, 'ok',
    [{ tekst: 'Otwórz', fn: () => api('otworz_plik', r.plik) }, { tekst: 'Folder', fn: () => api('otworz_folder', r.plik) }], 8000);
}

// ------------------------------------------------------------------ baza materiałów, silników i chłodnic
const RODZAJE = { materials: 'Materiały', engines: 'Silniki', radiators: 'Chłodnice' };
const NAZWA_POJ = { materials: 'materiał', engines: 'silnik', radiators: 'chłodnica' };
function opisUrzadzenia(rodzaj, u) {
  const pliki = u.pliki ? ` · ${Object.keys(u.pliki).length === 2 ? 'obraz + karta' : u.pliki.obraz ? 'obraz' : 'karta'}` : '';
  if (rodzaj === 'materials') return (u.type === 'Rw / Ctr' ? `Rw ${lc(u.rw)} / Ctr ${lc(u.ctr)} dB` : `widmo ${(u.spec || []).map((x) => lc(x)).join(' / ')} dB`) + (+u.d ? ` · ${lc(u.d)} mm` : '') + pliki;
  if (rodzaj === 'engines') return `${esc(u.genset || '')}${u.genset ? ' · ' : ''}LwA ${lp(sumaDb(widmoA(u.lw || [], u.jedn)), 1)} dB(A)${u.jedn === 'dB' ? ' (widmo w dB)' : ''}${pliki}`;
  return `LwA ${lp(sumaDb(widmoA(u.lw || [], u.jedn)), 1)} dB(A)${u.jedn === 'dBA' ? '' : ' (widmo w dB)'}${pliki}`;
}

function rysujBaze() {
  const el = $('#widok-baza');
  el.className = 'widok baza-widok';
  const lista = Object.entries(BAZA[baza.rodzaj]).filter(([n]) => n.toLowerCase().includes(baza.szukaj.toLowerCase())).sort(([a], [b]) => a.localeCompare(b, 'pl'));
  el.innerHTML = `<div class="karta lista-urzadzen">
      <div class="gora-listy">
        <div class="segment pelny" id="b-rodzaj">${Object.entries(RODZAJE).map(([k, t]) => `<button data-v="${k}" class="${baza.rodzaj === k ? 'akt' : ''}">${t} <span class="podp">${Object.keys(BAZA[k]).length}</span></button>`).join('')}</div>
        <div class="z-jednostka"><input class="pole" id="b-szukaj" placeholder="Szukaj…" value="${esc(baza.szukaj)}" style="font-family:var(--font-body)"><span class="jedn">${ik('szukaj')}</span></div>
        <button class="btn btn-maly" id="b-nowe">${ik('plus')} Nowa pozycja</button>
      </div>
      <div class="przewijana-lista">${lista.map(([n, u]) => { const z = BAZA.zrodla?.[baza.rodzaj]?.[n] || 'globalna'; return `<button class="urz ${baza.wybrany === n ? 'akt' : ''} ${z !== 'globalna' ? 'lokalna' : ''}" data-urz="${esc(n)}" title="${z === 'wlasna' ? 'Wpis własny - nie ma go w bazie globalnej' : z === 'zmieniona' ? 'Zmieniona przez Ciebie wersja pozycji z bazy globalnej' : 'Z bazy globalnej Diamel'}"><b>${esc(n)}</b><span>${z === 'wlasna' ? '<em>własny</em> · ' : z === 'zmieniona' ? '<em>zmieniony</em> · ' : ''}${opisUrzadzenia(baza.rodzaj, u)}</span></button>`; }).join('') || '<p class="podp" style="padding:12px">Brak pozycji.</p>'}</div>
    </div>
    <div class="edytor" id="edytor"></div>`;
  $$('#b-rodzaj button').forEach((b) => b.onclick = () => { baza.rodzaj = b.dataset.v; baza.wybrany = null; rysujBaze(); });
  $('#b-szukaj').oninput = (e) => { baza.szukaj = e.target.value; const poz = e.target.selectionStart; rysujBaze(); const s = $('#b-szukaj'); s.focus(); s.setSelectionRange(poz, poz); };
  $('#b-nowe').onclick = () => { baza.wybrany = ''; rysujBaze(); $('#e-nazwa')?.focus(); };
  $$('[data-urz]').forEach((b) => b.onclick = () => { baza.wybrany = b.dataset.urz; rysujBaze(); });
  rysujEdytor();
}

function rysujEdytor() {
  const el = $('#edytor');
  const nowy = baza.wybrany === '';
  if (baza.wybrany === null) {
    el.innerHTML = `<div class="karta"><h3>${ik('baza')} Baza materiałów, silników i chłodnic <span class="podp" style="font-weight:400">v${esc(BAZA.wersja)}</span></h3>
      <p class="opis">Twoja lokalna baza - pozycje z bazy globalnej Diamel i Twoje własne. Wybierz pozycję z listy albo dodaj nową.</p>
      <p class="opis">Baza globalna jest pobierana w całości z serwera przy starcie programu. Wpisy oznaczone <b style="color:var(--zolty)">●</b> są Twoje - aktualizacja ich nie zmienia.</p>
      <div class="wiersz"><span class="et">Baza globalna Diamel v${esc(BAZA.wersja)}${AKT_BAZY?.ok ? (AKT_BAZY.nowsza ? ` - <b style="color:var(--zolty-tekst)">na serwerze v${esc(AKT_BAZY.wersja)}</b>` : ' - aktualna') : ''}</span>
        <button class="btn btn-maly ${AKT_BAZY?.nowsza ? 'btn-hr' : ''}" id="b-akt">${ik('pobierz')} Pobierz z serwera</button></div>
      <div class="wiersz"><span class="et">Brakuje materiału, silnika albo chłodnicy?<small>Wyślij prośbę - po dodaniu trafi do wszystkich użytkowników.</small></span><button class="btn btn-maly" id="b-zglos">${ik('wyslij')} Zgłoś do bazy</button></div>
      <div class="wiersz"><span class="et">Pliki bazy (kopie przed aktualizacją są obok)</span><button class="btn btn-maly" id="b-folder">${ik('folder')} Otwórz folder</button></div></div>`;
    $('#b-akt').onclick = () => sprawdzBaze(true);
    $('#b-zglos').onclick = () => { zglos.rodzaj = 'sprzet'; zglos.kategoria = { materials: 'Materiał', engines: 'Silnik', radiators: 'Chłodnica' }[baza.rodzaj]; pokaz('zglos'); };
    $('#b-folder').onclick = () => api('folder_bazy');
    return;
  }
  const u = nowy ? { type: 'Widmo pasmowe', jedn: baza.rodzaj === 'engines' ? 'dBA' : 'dB' } : BAZA[baza.rodzaj][baza.wybrany] || {};
  const zrodlo = nowy ? 'wlasna' : BAZA.zrodla?.[baza.rodzaj]?.[baza.wybrany] || 'globalna';
  const pola = INFO.pola_urzadzen[baza.rodzaj];
  const wartosc = (k) => (u[k] === undefined || u[k] === null ? '' : u[k]);
  const rwCtr = baza.rodzaj === 'materials' && u.type === 'Rw / Ctr';
  el.innerHTML = `<div class="karta">
    <div class="naglowek-karty"><h3>${ik(baza.rodzaj === 'materials' ? 'warstwy' : 'silnik')} ${nowy ? `Nowy ${NAZWA_POJ[baza.rodzaj]}` : esc(baza.wybrany)}</h3>
      <span class="wartosci">${Object.entries(u.pliki || {}).map(([rola, p]) => `<button class="btn btn-maly" data-plik-bazy="${esc(p.id)}" title="${esc(p.nazwa)}">${ik(rola === 'obraz' ? 'obraz' : 'dokument')} ${rola === 'obraz' ? 'Obraz' : 'Karta katalogowa'}</button>`).join('')}
      ${!nowy ? `<button class="btn btn-maly" id="e-uzyj">${ik('dalej')} Użyj w projekcie</button>` : ''}</span></div>
    <div class="pola p2">
      <div class="p-pole szer2"><div class="podp" style="margin-bottom:8px">${zrodlo === 'globalna' ? 'Pozycja z bazy globalnej Diamel. Zapisanie zmian utworzy Twoją wersję (●) - baza globalna zostaje bez zmian.' : zrodlo === 'zmieniona' ? '● Twoja wersja pozycji z bazy globalnej - ma pierwszeństwo przed wersją z bazy.' : '● Wpis własny - nie ma go w bazie globalnej. Możesz go zgłosić do dodania (Zgłoszenia).'}</div>
        <label class="etykieta" for="e-nazwa">Nazwa</label><input class="pole" id="e-nazwa" value="${esc(nowy ? '' : baza.wybrany)}"></div>
      ${pola.map(([k, opis, jedn, rodzaj]) => {
        if (rodzaj === 'widmo') {
          if (k === 'spec' && rwCtr) return '';
          const w = wartosc(k) || [];
          return `<div class="p-pole szer2"><label class="etykieta">${opis}${jedn ? ` [${jedn}]` : ''}</label><div class="widmo-siatka mala">${INFO.pasma.map((p) => `<span class="pasmo">${p}</span>`).join('')}${INFO.pasma.map((_, i) => `<input class="pole" data-e-w="${k}" data-i="${i}" value="${esc(w[i] ?? '')}" inputmode="decimal">`).join('')}</div></div>`;
        }
        if (rodzaj === 'jedn') return `<div class="p-pole"><label class="etykieta">${opis}</label><select class="pole" data-e="${k}">${[['dBA', 'dB(A) - po korekcie A'], ['dB', 'dB - bez korekty A']].map(([v, t]) => `<option value="${v}" ${wartosc(k) === v ? 'selected' : ''}>${t}</option>`).join('')}</select></div>`;
        if (rodzaj === 'typ') return `<div class="p-pole"><label class="etykieta">${opis}</label><select class="pole" data-e="${k}" id="e-typ">${INFO.typy_materialu.map((t) => `<option ${wartosc(k) === t ? 'selected' : ''}>${t}</option>`).join('')}</select></div>`;
        if (rodzaj === true) return (['rw', 'ctr'].includes(k) && !rwCtr) ? '' : `<div class="p-pole"><label class="etykieta">${opis}</label><div class="z-jednostka"><input class="pole" data-e="${k}" value="${esc(wartosc(k) === 0 ? '' : wartosc(k))}" inputmode="decimal"><span class="jedn">${jedn}</span></div></div>`;
        return `<div class="p-pole"><label class="etykieta">${opis}</label><input class="pole" data-e="${k}" value="${esc(wartosc(k))}"></div>`;
      }).join('')}
    </div>
    <div class="przyciski-edytora">
      <button class="btn btn-hr" id="e-zapisz">${ik('zapisz')} Zapisz w bazie</button>
      ${zrodlo === 'wlasna' && !nowy ? `<button class="btn btn-czerwony" id="e-usun">${ik('kosz')} Usuń</button>` : zrodlo === 'zmieniona' ? `<button class="btn" id="e-usun">${ik('odswiez')} Przywróć z bazy globalnej</button>` : ''}
      <button class="btn btn-bez" id="e-anuluj">Anuluj</button>
    </div></div>`;
  const zbierz = () => {
    const dane = Object.fromEntries($$('[data-e]', el).map((x) => [x.dataset.e, x.value]));
    for (const x of $$('[data-e-w]', el)) { (dane[x.dataset.eW] ||= Array(8).fill(0))[+x.dataset.i] = x.value; }
    return { ...u, ...dane };
  };
  $('#e-typ')?.addEventListener('change', () => { const d = zbierz(); BAZA_EDYCJA = d; rysujEdytorZ(d); });
  $('#e-anuluj').onclick = () => { baza.wybrany = null; rysujBaze(); };
  $('#e-zapisz').onclick = async () => {
    const nazwa = $('#e-nazwa').value.trim();
    const r = await api('zapisz_urzadzenie', baza.rodzaj, nowy ? '' : baza.wybrany, nazwa, zbierz());
    if (!r?.ok) return toast(esc(r?.blad || 'Nie udało się zapisać.'), 'blad');
    BAZA = r.baza; baza.wybrany = nazwa; rysujBaze();
    toast(`Zapisano w bazie: <b>${esc(nazwa)}</b>`, 'ok', null, 2500);
  };
  $('#e-usun')?.addEventListener('click', async () => {
    const przywroc = zrodlo === 'zmieniona';
    const w = await modal({ tytul: `${ik(przywroc ? 'odswiez' : 'kosz')} ${przywroc ? 'Przywrócić wersję z bazy globalnej?' : 'Usunąć z bazy?'}`,
      tresc: `<p><b>${esc(baza.wybrany)}</b> - ${przywroc ? 'Twoje zmiany zostaną usunięte, pozycja wróci do wartości z bazy globalnej.' : 'wpis własny zostanie usunięty.'} Zapisane projekty zachowują swoje wartości.</p>`,
      przyciski: [{ tekst: 'Anuluj', wartosc: null }, { tekst: przywroc ? 'Przywróć' : 'Usuń', klasa: przywroc ? 'btn-hr' : 'btn-czerwony', wartosc: 'tak' }] });
    if (w !== 'tak') return;
    const r = await api('usun_urzadzenie', baza.rodzaj, baza.wybrany);
    BAZA = r.baza; if (!przywroc) baza.wybrany = null; rysujBaze();
  });
  $('#e-uzyj')?.addEventListener('click', () => uzyjZBazy(baza.rodzaj, baza.wybrany));
  $$('[data-plik-bazy]', el).forEach((b) => b.onclick = async () => {
    b.disabled = true;
    const r = await api('otworz_plik_bazy', b.dataset.plikBazy);
    b.disabled = false;
    if (!r?.ok) toast(esc(r?.blad || 'Nie udało się otworzyć pliku.'), r?.zaloguj ? '' : 'blad', r?.zaloguj ? { tekst: 'Zaloguj', fn: () => ekranLogowania() } : null, 7000);
  });
}
let BAZA_EDYCJA = null;
/** Zmiana rodzaju danych materiału w edytorze (widmo / Rw-Ctr) - pola przerysowane z bieżącymi wartościami. */
function rysujEdytorZ(dane) {
  const n = baza.wybrany, poprzednie = BAZA[baza.rodzaj][n];
  const nazwa = $('#e-nazwa').value;
  BAZA[baza.rodzaj][n] = dane;
  rysujEdytor();
  if (poprzednie === undefined) delete BAZA[baza.rodzaj][n]; else BAZA[baza.rodzaj][n] = poprzednie;
  $('#e-nazwa').value = nazwa;
}

async function uzyjZBazy(rodzaj, nazwa) {
  if (rodzaj === 'engines') { pokaz('obliczenia'); zastosujSilnik(nazwa); return; }
  if (rodzaj === 'radiators') { pokaz('obliczenia'); zastosujChlodnice(nazwa); return; }
  const k = await modal({ tytul: `${ik('plus')} Dodaj warstwę „${esc(nazwa)}”`, tresc: `<p class="podp">Do którego modułu obudowy?</p><div class="lista-ostatnich">${INFO.moduly.map((m) => `
      <button data-wybor="${m}"><b>${MODULY[m]}</b><span>${P.modules_active[m] ? `warstwy: ${(P.layers[m] || []).length}` : 'wyłączony - zostanie włączony'}</span></button>`).join('')}</div>`,
    przyciski: [{ tekst: 'Anuluj', wartosc: null }] });
  if (!k) return;
  P.modules_active[k] = true;
  P.layers[k].push(warstwaZBazy(nazwa));
  zakladka = k;
  zmiana();
  pokaz('obliczenia');
  toast(`Dodano warstwę <b>${esc(nazwa)}</b> do modułu ${MODULY[k]}.`, 'ok', null, 3000);
}

// ------------------------------------------------------------------ ustawienia
function przelacznik(id, wl) { return `<button class="przelacznik ${wl ? 'wl' : ''}" id="${id}" role="switch" aria-checked="${wl}"></button>`; }

function rysujUstawienia() {
  const el = $('#widok-ustawienia');
  const pr = INFO.program;
  const seg = (id, opcje, wartosc) => `<div class="segment" id="${id}">${opcje.map(([v, t]) => `<button data-v="${v}" class="${wartosc === v ? 'akt' : ''}">${t}</button>`).join('')}</div>`;
  el.innerHTML = `<div class="siatka k2">
    <section class="karta"><h3>${ik('obraz')} Wygląd</h3><p class="opis">Motyw możesz też zmienić w menu z lewej</p>
      <div class="wiersz"><span class="et">Motyw</span>${seg('u-motyw', [['jasny', 'Jasny'], ['ciemny', 'Ciemny'], ['system', 'Jak w Windows']], pr.motyw)}</div>
      <div class="wiersz"><span class="et">Wielkość tekstu</span>${seg('u-rozmiar', [['maly', 'Mała'], ['normalny', 'Normalna'], ['duzy', 'Duża']], pr.rozmiar)}</div></section>
    <section class="karta"><h3>${ik('word')} Raporty</h3><p class="opis">Raport PDF / DOCX z obliczeń</p>
      <div class="wiersz"><span class="et">Otwórz raport po zapisaniu</span>${przelacznik('u-otworz', pr.otworz_raport)}</div>
      <div class="wiersz"><span class="et">Karty producentów w raporcie PDF<small>pliki PDF silnika, chłodnicy, modułów i wydechu wstawione po ich rozdziałach</small></span>${przelacznik('u-pdf-zal', pr.pdf_zalaczniki)}</div>
      <div class="wiersz"><span class="et">${KONTO.szablon ? 'Własny szablon (zastępuje papier firmowy)' : 'Szablon z nagłówkiem i stopką'}<small>${pr.szablon ? esc(pr.szablon) : KONTO.szablon ? `brak - raporty na papierze firmowym ${esc(KONTO.firma)} (z konta)` : 'brak - stopka programu (numer projektu, numer strony)'}</small></span>
        <span class="wartosci">${pr.szablon ? '<button class="btn btn-maly btn-bez" id="u-szablon-usun">Usuń</button>' : ''}<button class="btn btn-maly" id="u-szablon">${ik('folder')} Wybierz…</button></span></div>
      <p class="podp" style="margin:8px 0 0">Szablon to zwykły dokument Word (.docx) z Twoim nagłówkiem i stopką. Program bierze z niego nagłówek, stopkę, marginesy i rozmiar strony - treść szablonu pomija. PDF zapisuje Microsoft Word, więc stopka jest taka sama w PDF i DOCX. W nagłówku i stopce szablonu możesz wpisać pola, które program uzupełni: <b>{PROJEKT}</b> <b>{KLIENT}</b> <b>{AGREGAT}</b> <b>{SILNIK}</b> <b>{AUTOR}</b> <b>{DATA}</b> <b>{TYTUL}</b>.</p></section>
    <section class="karta"><h3>${ik('nowy')} Nowy projekt</h3><p class="opis">Wartości wpisywane w nowym projekcie</p>
      <div class="wiersz"><span class="et">Przygotował (inicjały)${KONTO.inicjaly ? `<small>z konta: ${esc(KONTO.inicjaly)}</small>` : ''}</span><input class="pole waski" id="u-autor" value="${esc(pr.autor)}"></div>
      <div class="wiersz"><span class="et">Język raportu</span>${seg('u-jezyk', [['English', 'Angielski'], ['Polski', 'Polski']], pr.jezyk)}</div></section>
    <section class="karta"><h3>${ik('baza')} Konto i dane</h3><p class="opis">Konto Diamel i baza materiałów, silników i chłodnic z serwera diamel.pl</p>
      <div class="wiersz"><span class="et">Zalogowano${KONTO.zalogowany ? `<small>${esc(KONTO.nazwa)}${KONTO.inicjaly ? ` · inicjały w raportach: ${esc(KONTO.inicjaly)}` : ''}</small>` : ''}</span><b>${KONTO.zalogowany ? esc(KONTO.login) : 'nie'}</b></div>
      <div class="wiersz"><span class="et">Papier firmowy w raportach${KONTO.zalogowany ? `<small>${KONTO.szablon ? `${esc(KONTO.firma)} · ${esc(KONTO.szablon)}` : KONTO.firma ? `${esc(KONTO.firma)} nie ma jeszcze szablonu w panelu` : 'konto bez przypisanej firmy'}</small>` : ''}</span>${KONTO.szablon ? '<span class="status ok">aktywny</span>' : '<span class="status szary">brak</span>'}</div>
      <div class="wiersz"><span class="et">Baza</span><span class="mono">v${esc(BAZA.wersja)}</span></div>
      <div class="wiersz"><span class="et"></span><span class="wartosci"><button class="btn btn-maly" id="u-baza">${ik('odswiez')} Sprawdź bazę</button>${KONTO.zalogowany ? '<button class="btn btn-maly btn-czerwony" id="u-wyloguj">Wyloguj</button>' : '<button class="btn btn-maly btn-hr" id="u-zaloguj">Zaloguj</button>'}</span></div></section>
    <section class="karta" style="grid-column: span 2"><h3>${ik('pobierz')} O programie</h3><p class="opis">Kalkulator Akustyki v${esc(INFO.wersja)} · ${esc(INFO.data)} · © Mateusz Kujawa, DIAMEL</p>
      <div class="wiersz"><span class="et">${AKT?.do_instalacji ? `${AKT.wycofana ? 'Zalecana (starsza) wersja' : 'Dostępna wersja'} <b>${esc(AKT.wersja)}</b>` : 'Aktualizacje'}</span>
        <span class="wartosci"><button class="btn btn-maly" id="u-zmiany">Historia zmian</button><button class="btn btn-maly" id="u-www">Strona programu</button><button class="btn btn-maly ${AKT?.do_instalacji ? 'btn-hr' : ''}" id="u-akt">${AKT?.do_instalacji ? 'Zainstaluj' : 'Sprawdź aktualizacje'}</button></span></div></section>
  </div>`;
  $('#u-autor').onchange = (e) => zmienProgram('autor', e.target.value.trim());
  const naSeg = (id, klucz) => $$(`#${id} button`).forEach((b) => b.onclick = async () => { await zmienProgram(klucz, b.dataset.v); rysujUstawienia(); });
  naSeg('u-jezyk', 'jezyk'); naSeg('u-motyw', 'motyw'); naSeg('u-rozmiar', 'rozmiar');
  $('#u-pdf-zal').onclick = async () => { await zmienProgram('pdf_zalaczniki', !INFO.program.pdf_zalaczniki); rysujUstawienia(); };
  $('#u-otworz').onclick = async () => { await zmienProgram('otworz_raport', !INFO.program.otworz_raport); rysujUstawienia(); };
  $('#u-wyloguj')?.addEventListener('click', async () => {
    const w = await modal({ tytul: 'Wylogować?', tresc: '<p>Zapamiętany dostęp na tym komputerze zostanie usunięty, a raporty będą miały stopkę programu. Przy następnym uruchomieniu program zapyta o logowanie.</p>', przyciski: [{ tekst: 'Anuluj', wartosc: null, domyslny: true }, { tekst: 'Wyloguj', klasa: 'btn-czerwony', wartosc: 'tak' }] });
    if (w !== 'tak') return;
    KONTO = (await api('konto_wyloguj')).stan;
    poZmianieKonta();
    ekranLogowania();
  });
  $('#u-zaloguj')?.addEventListener('click', () => ekranLogowania());
  $('#u-szablon').onclick = async () => { INFO.program = await api('wybierz_szablon'); rysujUstawienia(); };
  $('#u-szablon-usun')?.addEventListener('click', async () => { await zmienProgram('szablon', ''); rysujUstawienia(); });
  $('#u-baza').onclick = () => sprawdzBaze(true);
  $('#u-zmiany').onclick = () => modal({ tytul: 'Historia zmian', szeroki: true, tresc: `<div style="white-space:pre-wrap; font-size:0.86rem; max-height:60vh; overflow:auto">${esc(INFO.changelog)}</div>`, przyciski: [{ tekst: 'Zamknij', wartosc: null }] });
  $('#u-www').onclick = () => api('otworz_strone', 'https://github.com/Kuj0n/Diamel-Akustyka/releases');
  $('#u-akt').onclick = () => (AKT?.do_instalacji ? oknoAktualizacji() : sprawdzAktualizacje(true));
}

// ------------------------------------------------------------------ zgłoszenia
const zglos = { rodzaj: 'sprzet', kategoria: 'Materiał', email: '', nazwa: '', opis: '', projekt: true, zalaczniki: [] };
function rysujZglos() {
  const el = $('#widok-zglos');
  const sprzet = zglos.rodzaj === 'sprzet';
  el.innerHTML = `<div class="karta formularz">
    <div class="segment pelny" id="z-rodzaj"><button data-v="sprzet" class="${sprzet ? 'akt' : ''}">${ik('baza')} Prośba o dodanie do bazy</button><button data-v="blad" class="${!sprzet ? 'akt' : ''}">${ik('uwaga')} Błąd programu</button></div>
    <p class="opis" style="margin:0">${sprzet ? 'Brakuje materiału, silnika albo chłodnicy? Podaj nazwę i dołącz kartę katalogową z widmem / izolacyjnością - po weryfikacji trafi do bazy globalnej i do wszystkich użytkowników.' : 'Opisz, co się stało. Możesz dołączyć bieżący projekt - ułatwi to odtworzenie błędu.'}</p>
    <div class="pola p2">
      <div class="p-pole"><label class="etykieta" for="z-email">Twój e-mail (do odpowiedzi)</label><input class="pole" id="z-email" type="email" value="${esc(zglos.email)}"></div>
      ${sprzet ? `<div class="p-pole"><label class="etykieta">Kategoria</label><div class="segment pelny" id="z-kat">${['Materiał', 'Silnik', 'Chłodnica'].map((k) => `<button data-v="${k}" class="${zglos.kategoria === k ? 'akt' : ''}">${k}</button>`).join('')}</div></div>
        <div class="p-pole szer2"><label class="etykieta" for="z-nazwa">Producent i nazwa / model</label><input class="pole" id="z-nazwa" value="${esc(zglos.nazwa)}"></div>` : ''}
    </div>
    <div><label class="etykieta" for="z-opis">${sprzet ? 'Parametry (widmo, izolacyjność, grubość), link do karty katalogowej' : 'Opis problemu'}</label><textarea class="pole" id="z-opis" rows="7">${esc(zglos.opis)}</textarea></div>
    <div><span class="etykieta">Załączniki${sprzet ? ' (karta katalogowa, protokół badań)' : ''} - razem do 5 MB</span>
      <div class="lista-zalacznikow">${zglos.zalaczniki.map((z, i) => `<span class="znacznik akt">${ik('plik')} ${esc(z.nazwa)} <span class="podp">${z.kb} kB</span><button class="btn btn-bez btn-ikona" data-usun-zal="${i}" title="Usuń">${ik('x')}</button></span>`).join('')}
        <button class="btn btn-maly" id="z-dodaj-zal">${ik('plus')} Dodaj pliki…</button></div></div>
    ${!sprzet ? `<label class="zaznacz"><input type="checkbox" id="z-projekt" ${zglos.projekt ? 'checked' : ''}> Dołącz bieżący projekt (${esc(P.meta.project_no || 'bez nazwy')})</label>` : ''}
    <div><button class="btn btn-hr" id="z-wyslij">${ik('wyslij')} Wyślij</button></div>
  </div>`;
  const zapamietaj = () => { zglos.email = $('#z-email').value; zglos.opis = $('#z-opis').value; if ($('#z-nazwa')) zglos.nazwa = $('#z-nazwa').value; if ($('#z-projekt')) zglos.projekt = $('#z-projekt').checked; };
  $$('#z-rodzaj button').forEach((b) => b.onclick = () => { zapamietaj(); zglos.rodzaj = b.dataset.v; rysujZglos(); });
  $$('#z-kat button').forEach((b) => b.onclick = () => { zapamietaj(); zglos.kategoria = b.dataset.v; rysujZglos(); });
  $('#z-dodaj-zal').onclick = async () => {
    zapamietaj();
    const nowe = (await api('wybierz_zalaczniki')) || [];
    zglos.zalaczniki = [...zglos.zalaczniki, ...nowe.filter((n) => !zglos.zalaczniki.some((z) => z.sciezka === n.sciezka))];
    const kb = zglos.zalaczniki.reduce((s, z) => s + z.kb, 0);
    if (kb > 5 * 1024) toast(`Załączniki mają razem ${lp(kb / 1024, 1)} MB - limit to 5 MB.`, 'blad');
    rysujZglos();
  };
  $$('[data-usun-zal]').forEach((b) => b.onclick = () => { zapamietaj(); zglos.zalaczniki.splice(+b.dataset.usunZal, 1); rysujZglos(); });
  $('#z-wyslij').onclick = async () => {
    zapamietaj();
    if (!/^\S+@\S+\.\S+$/.test(zglos.email.trim())) { toast('Podaj poprawny adres e-mail.', 'blad'); return $('#z-email').focus(); }
    if (sprzet && !zglos.nazwa.trim()) { toast('Podaj nazwę.', 'blad'); return $('#z-nazwa').focus(); }
    if (!sprzet && !zglos.opis.trim()) { toast('Opisz problem.', 'blad'); return $('#z-opis').focus(); }
    const b = $('#z-wyslij'); b.disabled = true; b.textContent = 'Wysyłanie…';
    const pliki = zglos.zalaczniki.map((z) => z.sciezka);
    const r = sprzet ? await api('zglos', zglos.email.trim(), zglos.kategoria, zglos.nazwa.trim(), zglos.opis.trim(), null, pliki)
      : await api('zglos', zglos.email.trim(), 'Błąd programu', `Zgłoszenie błędu Kalkulatora v${INFO.wersja}`, zglos.opis.trim(), zglos.projekt ? P : null, pliki);
    b.disabled = false; b.innerHTML = `${ik('wyslij')} Wyślij`;
    if (!r?.ok) return toast(`Nie udało się wysłać: ${esc(r?.blad || '')}`, 'blad');
    toast('Zgłoszenie wysłane - odpowiemy na podany adres e-mail. Dziękujemy!', 'ok', null, 7000);
    zglos.nazwa = ''; zglos.opis = ''; zglos.zalaczniki = []; rysujZglos();
  };
}

// ------------------------------------------------------------------ skróty klawiszowe
document.addEventListener('keydown', (e) => {
  if (modalOtwarty) { if (e.key === 'Escape') modalOtwarty.zamknij(null); return; }
  const kl = e.key.toLowerCase();
  if (e.ctrlKey && kl === 's') { e.preventDefault(); return zapiszProjekt(e.shiftKey); }
  if (e.ctrlKey && kl === 'o') { e.preventDefault(); return otworzProjekt(); }
  if (e.ctrlKey && kl === 'n') { e.preventDefault(); return nowyProjekt(); }
  if (e.ctrlKey && kl === 'g') { e.preventDefault(); if (widok === 'obliczenia') generujRaport(e.shiftKey ? 'docx' : 'pdf'); }
});
