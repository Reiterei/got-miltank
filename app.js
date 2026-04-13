// ── CONFIGURATION ────────────────────────────────────
// Replace these with your own Google Sheets published CSV URLs
// See README for instructions on how to set these up
var CONFIG = {
  CARDS_SHEET_URL:   'https://docs.google.com/spreadsheets/d/e/2PACX-1vRdF01jJlXmqVYPnnUeZcH-FUkMfk0lvD6WFw0qS-c0UVr8yUEsEgdXtLHWq1TcWlh6SU7LZkDWfxmP/pub?gid=1984372806&single=true&output=csv',
  SETS_SHEET_URL:    'https://docs.google.com/spreadsheets/d/e/2PACX-1vStKDQ-Z6Hm5eH5v8qAxuq6w2jGW_jnyeaKcqrgm1HKGCKi4hkqg3Voy4_xOfovthzfM-qduwqVDE6s/pub?gid=727524711&single=true&output=csv',
  POKEMON_SHEET_URL: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRDzKb6cjvjhYlbU_2-23WzMJ9t2tWOz9n0TePOtiWtX0hZ7i88HnptGvXh7eLz_ANUg_cNv5zFawQ_/pub?gid=1479465458&single=true&output=csv',
};

// ── DATA (populated from Google Sheets on load) ───────
var SET_DATA      = {};
var SET_CARD_DATA = {};
var KNOWN_SETS    = [];
var POKEMON_LIST  = [];
var BASE_SET_CARDS = [];
var _dexLabels = [];
var BLANK_SYMBOL_URI = 'data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32"><circle cx="16" cy="16" r="14" fill="%23ddd"/></svg>';
var CORNER_DOTS = '<span class="corner-dot tl"></span><span class="corner-dot tr"></span>'
               + '<span class="corner-dot bl"></span><span class="corner-dot br"></span>';
var BORDER_TYPES = ['border-dots','border-solid','border-dotted','border-none'];

function getSymbolUrl(setName) {
  var info = SET_DATA[setName];
  var symbol = info && info.symbol ? info.symbol : 'Base_Set.png';
  if (symbol.indexOf('http') === 0 || symbol.indexOf('//') === 0) return symbol;
  return 'symbols/' + symbol;
}

function getScrydexImageUrl(setName, cardNum, customImageUrl) {
  if (customImageUrl) return customImageUrl;
  var info = SET_DATA[setName];
  if (!info || !info.scrydexId || !cardNum) return null;
  // Card number may be "41/102" or "041/102" — scrydex wants just the numeric part before the slash
  var num = String(cardNum).split('/')[0].replace(/^0+/, '') || cardNum.split('/')[0];
  return 'https://images.scrydex.com/pokemon/' + info.scrydexId + '-' + num + '/medium';
}

// Returns the card object with the given id, or undefined.
function findCard(id) {
  return cards.find(function(c) { return c.id === id; });
}

// Returns the display string for the currently selected Pokémon filter.
function pokemonDisplayValue() {
  if (!pokemonFilter) return 'Pokémon Species';
  var dexIdx = parseInt(pokemonFilter.replace('#', ''), 10) - 1;
  return (dexIdx >= 0 && POKEMON_LIST[dexIdx])
    ? pokemonFilter + ' - ' + POKEMON_LIST[dexIdx]
    : pokemonFilter;
}


// ── PAPER CONFIG ─────────────────────────────────────
var PAPER_DATA = {
  'Letter-L':  { css:'letter landscape',  pageW:11.0,  pageH:8.5,   margin:0.5,  cols:3, rows:2 },
  'Legal-L':   { css:'legal landscape',   pageW:14.0,  pageH:8.5,   margin:0.5,  cols:4, rows:2 },
  'A4-L':      { css:'A4 landscape',      pageW:11.69, pageH:8.27,  margin:0.35, cols:4, rows:2 },
};

var CARD_W_PX    = 240;
var CARD_H_PX    = 336;
var DPI          = 96;
var cards        = [];
var cardIdCtr    = 0;
var pokemonFilter = ''; // '' = All
var includeCameo = false;
var includeSerials = false;
var currentPaper = 'Letter-L';
var currentBorderType = 'dots';
var hideQRCard = false;
var showCardImages = false;
var collapsedIds = {};

// ── POKEMON FILTER ───────────────────────────────────
var activePokemonPickerId = null;

// Single global mousedown handler closes any open picker when clicking outside it
document.addEventListener('mousedown', function(e) {
  if (activePokemonPickerId !== null) {
    var pDrop  = document.getElementById('poke-picker-drop-global');
    var pInput = document.getElementById('poke-picker-input-global');
    if (!((pDrop && pDrop.contains(e.target)) || e.target === pInput)) closePokemonPicker();
  }
  if (activePickerId !== null) {
    var cDrop  = document.getElementById('picker-drop-'  + activePickerId);
    var cInput = document.getElementById('picker-input-' + activePickerId);
    if (!((cDrop && cDrop.contains(e.target)) || e.target === cInput)) closeCardPicker(activePickerId);
  }
  if (activeSetPickerId !== null) {
    var sDrop  = document.getElementById('set-picker-drop-'  + activeSetPickerId);
    var sInput = document.getElementById('set-picker-input-' + activeSetPickerId);
    if (!((sDrop && sDrop.contains(e.target)) || e.target === sInput)) closeSetPicker(activeSetPickerId);
  }
});

// ── GENERIC PICKER HELPERS ───────────────────────────
function buildPickerHTML(items, isSelectedFn, labelFn, handlerFn, valueFn, emptyMsg) {
  var opts = items.reduce(function(acc, item) {
    return acc + '<div class="card-picker-option' + (isSelectedFn(item) ? ' highlighted' : '') + '"'
      + ' data-value="' + esc(String(valueFn(item))) + '"'
      + ' onmousedown="' + handlerFn(item) + '">'
      + esc(labelFn(item)) + '</div>';
  }, '');
  return opts || '<div class="card-picker-option no-results">' + emptyMsg + '</div>';
}

function openPickerEl(drop, input, html) {
  if (!drop || !input) return;
  // First open: keep readonly so keyboard doesn't appear; second click removes it
  input.value = '';
  drop.innerHTML = html;
  drop.classList.add('open');
  setTimeout(function() {
    var sel = drop.querySelector('.highlighted');
    if (sel) sel.scrollIntoView({ block: 'nearest' });
  }, 0);
}

function isMobileView() {
  return window.matchMedia('(max-width: 600px), (max-height: 500px) and (orientation: landscape)').matches;
}

function enablePickerTyping(input) {
  if (!input) return;
  input.removeAttribute('readonly');
  input.focus();
}

function closePickerEl(drop, input, restoreValue) {
  if (drop) drop.classList.remove('open');
  if (input) { input.value = restoreValue; input.setAttribute('readonly', ''); }
}

function buildPokemonOptions(filter) {
  var f = filter.toLowerCase();
  var items = [];
  for (var pi = 0; pi < POKEMON_LIST.length; pi++) {
    var label  = _dexLabels[pi];
    var dexNum = '#' + String(pi + 1).padStart(4, '0');
    if (!f || label.toLowerCase().indexOf(f) !== -1) items.push({ dexNum: dexNum, label: label });
  }
  return buildPickerHTML(items,
    function(it) { return it.dexNum === pokemonFilter; },
    function(it) { return it.label; },
    function(it) { return 'event.preventDefault();pickPokemon(event,\'' + it.dexNum.replace(/'/g, "\\'") + '\')'; },
    function(it) { return it.dexNum; },
    'No Pokémon found'
  );
}

function openPokemonPicker() {
  activePokemonPickerId = 'global';
  openPickerEl(
    document.getElementById('poke-picker-drop-global'),
    document.getElementById('poke-picker-input-global'),
    buildPokemonOptions('')
  );
}

function closePokemonPicker() {
  activePokemonPickerId = null;
  closePickerEl(
    document.getElementById('poke-picker-drop-global'),
    document.getElementById('poke-picker-input-global'),
    pokemonDisplayValue()
  );
}

function filterPokemonPicker(value) {
  var drop = document.getElementById('poke-picker-drop-global');
  if (!drop) return;
  drop.innerHTML = buildPokemonOptions(value);
  drop.classList.add('open');
}

function openPokemonPickerMousedown(e) {
  e.preventDefault();
  if (activePokemonPickerId === 'global') {
    var input = document.getElementById('poke-picker-input-global');
    if (isMobileView() && input && input.hasAttribute('readonly')) { enablePickerTyping(input); return; }
    closePokemonPicker(); return;
  }
  openPokemonPicker();
  if (!isMobileView()) enablePickerTyping(document.getElementById('poke-picker-input-global'));
}

var _pendingPokemonChange = null;

function pickPokemon(e, name) {
  e.preventDefault();
  if (name === pokemonFilter) { closePokemonPicker(); return; }
  if (cards.length > 0) {
    _pendingPokemonChange = name;
    closePokemonPicker();
    document.getElementById('pokemon-confirm-modal').classList.add('open');
    return;
  }
  applyPokemonChange(name);
}

function closeModal(id) {
  document.getElementById(id).classList.remove('open');
}

function confirmPokemonChange() {
  var pending = _pendingPokemonChange;
  _pendingPokemonChange = null;
  closeModal('pokemon-confirm-modal');
  if (pending !== null) {
    cards = [];
    applyPokemonChange(pending);
  }
}

function closePokemonConfirmModal() {
  closeModal('pokemon-confirm-modal');
  _pendingPokemonChange = null;
}

function applyPokemonChange(newFilter) {
  pokemonFilter = newFilter;
  cards.forEach(function(c) {
    if (pokemonFilter && c.set) {
      var scd = SET_CARD_DATA[c.set];
      if (scd) {
        var hasMatch = scd.some(function(card) { return matchesPokemonFilter(card); });
        if (!hasMatch) {
          c.set = ''; c.cardKey = -1; c.name = '';
          c.variants = [];
        }
      }
    }
  });
  closePokemonPicker();
  render();
}

// Shared arrow-key/enter/escape navigation for all dropdown pickers.
function pickerKeyNav(e, drop, onEnter, onEscape) {
  if (!drop || !drop.classList.contains('open')) return;
  var opts = drop.querySelectorAll('.card-picker-option:not(.no-results)');
  var highlighted = drop.querySelector('.card-picker-option.highlighted');
  var idx = -1;
  opts.forEach(function(o, i) { if (o === highlighted) idx = i; });
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    var next = idx < opts.length - 1 ? idx + 1 : 0;
    if (highlighted) highlighted.classList.remove('highlighted');
    if (opts[next]) { opts[next].classList.add('highlighted'); opts[next].scrollIntoView({block:'nearest'}); }
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    var prev = idx > 0 ? idx - 1 : opts.length - 1;
    if (highlighted) highlighted.classList.remove('highlighted');
    if (opts[prev]) { opts[prev].classList.add('highlighted'); opts[prev].scrollIntoView({block:'nearest'}); }
  } else if (e.key === 'Enter') {
    e.preventDefault();
    if (highlighted) onEnter(highlighted);
  } else if (e.key === 'Escape') {
    onEscape();
  }
}

function pokemonPickerKeydown(e) {
  pickerKeyNav(e,
    document.getElementById('poke-picker-drop-global'),
    function(h) { pickPokemon(e, h.dataset.value); },
    closePokemonPicker
  );
}

function cardPickerKeydown(e, cardId) {
  pickerKeyNav(e,
    document.getElementById('picker-drop-' + cardId),
    function(h) { pickCard(e, cardId, parseInt(h.dataset.value)); },
    function() { closeCardPicker(cardId); }
  );
}

function setPickerKeydown(e, cardId) {
  pickerKeyNav(e,
    document.getElementById('set-picker-drop-' + cardId),
    function(h) { pickSet(e, cardId, h.dataset.value); },
    function() { closeSetPicker(cardId); }
  );
}

function toggleCameo(checked) {
  includeCameo = checked;
  render();
}

function toggleSerials(checked) {
  includeSerials = checked;
  if (!checked) {
    cards.forEach(function(c) {
      if (c.cardKey < 0 || !c.set) return;
      var sclP = SET_CARD_DATA[c.set];
      if (!sclP || !sclP[c.cardKey]) return;
      var cardData = sclP[c.cardKey];
      var serials = cardData.serialVariants || [];
      if (!serials.length) return;
      var hasSerialSelected = serials.some(function(s) { return (c.variants || []).indexOf(s) !== -1; });
      var hasNonHolo = (c.variants || []).indexOf('Non-Holo') !== -1;
      var nonHoloAvailable = (cardData.variants || []).indexOf('Non-Holo') !== -1;
      if (hasSerialSelected && !hasNonHolo && nonHoloAvailable) {
        c.variants = (c.variants || []).concat(['Non-Holo']);
      }
    });
  }
  render();
}

function matchesPokemonFilter(card) {
  if (!pokemonFilter) return true;
  var f = pokemonFilter.toUpperCase();
  function matchesField(field) {
    return field && field.split(',').some(function(p) { return p.trim().toUpperCase() === f; });
  }
  return matchesField(card.primary) || (includeCameo && matchesField(card.cameo));
}

function setsForPokemon() {
  if (!pokemonFilter) return KNOWN_SETS;
  return KNOWN_SETS.filter(function(s) {
    var scd = SET_CARD_DATA[s];
    return scd && scd.some(function(c) { return matchesPokemonFilter(c); });
  });
}

function cardsForPokemon(setName) {
  var scd = SET_CARD_DATA[setName];
  if (!scd) return [];
  if (!pokemonFilter) return scd;
  return scd.filter(function(c) { return matchesPokemonFilter(c); });
}

function toggleCardCollapse(cardId, e) {
  if (e && e.target && (e.target.classList.contains('remove-btn'))) return;
  var form = document.getElementById('form-' + cardId);
  if (!form) return;
  var isCollapsed = form.classList.toggle('collapsed');
  if (isCollapsed) { collapsedIds[cardId] = true; } else { delete collapsedIds[cardId]; }
}

function setBorderType(type) {
  currentBorderType = type;
  var grids = document.querySelectorAll('.card-grid');
  grids.forEach(function(g) {
    BORDER_TYPES.forEach(function(c) { g.classList.remove(c); });
    g.classList.add('border-' + type);
  });
  saveState();
}

function toggleHideQR(checked) {
  hideQRCard = checked;
  renderPreview();
}

function toggleShowCardImages(checked) {
  showCardImages = checked;
  saveState();
  renderPreview();
}

function toggleSettings() {
  var panel = document.getElementById('settings-panel');
  if (panel) panel.classList.toggle('open');
}

function logoEasterEgg() {
  if (!pokemonFilter) return;
  var name = pokemonDisplayValue().split(' - ')[1] || '';
  if (name.toLowerCase() !== 'miltank') return;
  if (miltankPlaying) { stopMiltankPlayer(); return; }
  startMiltankPlayer();
}

var miltankPlaying = false;
var miltankVolume  = 10;
var _miltankAudio  = null;

function startMiltankPlayer() {
  miltankPlaying = true;
  if (!_miltankAudio) {
    _miltankAudio = new Audio('https://reiterei.github.io/got-miltank/i_like_miltank.mp3');
    _miltankAudio.loop = true;
  }
  _miltankAudio.volume = miltankVolume / 100;
  _miltankAudio.play();
  renderForms();
}

function stopMiltankPlayer() {
  miltankPlaying = false;
  if (_miltankAudio) { _miltankAudio.pause(); _miltankAudio.currentTime = 0; }
  renderForms();
}

function setMiltankVolume(val) {
  miltankVolume = parseInt(val, 10);
  if (_miltankAudio) _miltankAudio.volume = miltankVolume / 100;
}

function render() {
  renderForms();
  renderPreview();
}

function initApp() {
  hideLoadingScreen();
  restoreState();
  setPaper(currentPaper);
  renderForms();
}

function setPaper(name) {
  currentPaper = name;
  var p = PAPER_DATA[name];
  var el = document.getElementById('print-page-style');
  if (!el) { el = document.createElement('style'); el.id = 'print-page-style'; document.head.appendChild(el); }
  el.textContent = '@media print { @page { size: ' + p.css + '; margin: ' + p.margin + 'in; } }';
  renderPreview();
}

// ── GENERIC INPUT HANDLER ────────────────────────────
document.addEventListener('input', function(e) { handleFieldEvent(e); });

function handleFieldEvent(e) {
  var el    = e.target;
  var id    = parseInt(el.dataset.id);
  var field = el.dataset.field;
  if (!id || !field) return;
  var c = findCard(id);
  if (!c) return;
  c[field] = el.value;
  if (field === 'set') {
    c.variants  = [];
    c.cardKey   = -1;
    c.name      = '';
    render();
    return;
  }
  renderPreview();
}

// ── CARDS ────────────────────────────────────────────
function addCard(data) {
  data = data || {};
  var newCard = {
    id:              ++cardIdCtr,
    name:            data.name            || '',
    num:             data.num             || '',
    set:             data.set             || '',
    variants:        data.variants        || [],
    specialVariants: data.specialVariants || [],
    cardKey:         typeof data.cardKey !== 'undefined' ? data.cardKey : -1,
  };
  if (!newCard.set) {
    var availableSets = setsForPokemon();
    if (availableSets.length === 1) {
      newCard.set = availableSets[0];
      var availableCards = cardsForPokemon(newCard.set);
      if (availableCards.length === 1) {
        var allCards = SET_CARD_DATA[newCard.set] || [];
        var bc = availableCards[0];
        newCard.cardKey  = allCards.indexOf(bc);
        newCard.name     = bc.name;
        newCard.num      = bc.num || '';
        newCard.variants = [];
      }
    }
  }
  cards.push(newCard);
  render();
}

function removeCard(id) {
  cards = cards.filter(function(c) { return c.id !== id; });
  delete collapsedIds[id];
  render();
}

function duplicateCard(id) {
  var c = findCard(id);
  if (!c) return;
  cards.push({
    id:              ++cardIdCtr,
    name:            c.name,
    num:             c.num,
    set:             c.set,
    variants:        c.variants.slice(),
    specialVariants: c.specialVariants.slice(),
    cardKey:         c.cardKey,
  });
  render();
}


function toggleVariant(btn, variant, isSpecial) {
  var id = parseInt(btn.closest('.card-form').dataset.cardid);
  var c = findCard(id);
  if (!c) return;
  var arr = isSpecial ? c.specialVariants : c.variants;
  var idx = arr.indexOf(variant);
  if (idx === -1) { arr.push(variant); } else { arr.splice(idx, 1); }
  btn.classList.toggle('selected', arr.indexOf(variant) !== -1);
  updateVariantWarning(id, c);
  renderPreview();
}

function updateVariantWarning(cardId, c) {
  var form = document.getElementById('form-' + cardId);
  if (!form) return;
  var warning = form.querySelector('.variant-warning');
  var noneSelected = (c.variants || []).length === 0 && (c.specialVariants || []).length === 0;
  if (warning) {
    warning.style.display = noneSelected ? '' : 'none';
  }
}

function applySerialSubstitution(variants, cardData) {
  if (!includeSerials) return variants;
  var serials = Array.isArray(cardData.serialVariants) ? cardData.serialVariants : [];
  if (!serials.length) return variants;
  return serials.concat(variants.filter(function(v) { return v !== 'Non-Holo'; }));
}

function showClearConfirmModal() {
  document.getElementById('clear-confirm-modal').classList.add('open');
}
function closeClearConfirmModal() { closeModal('clear-confirm-modal'); }
function confirmClearAll() { closeClearConfirmModal(); clearAll(); }

function clearAll() {
  cards = [];
  localStorage.removeItem('gmState');
  render();
}

function confirmFillAllCards() {
  if (!pokemonFilter) return;
  if (cards.length > 0) {
    document.getElementById('fill-confirm-modal').classList.add('open');
    return;
  }
  fillAllCards();
}
function closeFillConfirmModal() { closeModal('fill-confirm-modal'); }
function confirmFillAll() { closeFillConfirmModal(); fillAllCards(); }

function fillAllCards() {
  if (!pokemonFilter) return;
  cards = [];
  cardIdCtr = 0;
  setsForPokemon().forEach(function(setName) {
    var allCards = SET_CARD_DATA[setName] || [];
    cardsForPokemon(setName).forEach(function(bc) {
      var variants = applySerialSubstitution(
        Array.isArray(bc.variants) ? bc.variants.slice() : [],
        bc
      );
      cards.push({
        id:              ++cardIdCtr,
        name:            bc.name,
        num:             bc.num || '',
        set:             setName,
        cardKey:         allCards.indexOf(bc),
        variants:        variants,
        specialVariants: Array.isArray(bc.specialVariants) ? bc.specialVariants.slice() : [],
      });
    });
  });
  render();
}

function saveState() {
  try {
    var state = {
      cards:             cards,
      pokemonFilter:     pokemonFilter,
      includeCameo:      includeCameo,
      includeSerials:    includeSerials,
      currentPaper:      currentPaper,
      currentBorderType: currentBorderType,
      hideQRCard:        hideQRCard,
      showCardImages:    showCardImages,
      cardIdCtr:         cardIdCtr,
    };
    localStorage.setItem('gmState', JSON.stringify(state));
  } catch(e) { /* storage unavailable */ }
}

function restoreState() {
  try {
    var raw = localStorage.getItem('gmState');
    if (!raw) return false;
    var state = JSON.parse(raw);
    cards             = state.cards             || [];
    pokemonFilter     = state.pokemonFilter     || '';
    includeCameo      = !!state.includeCameo;
    includeSerials    = !!state.includeSerials;
    currentPaper      = state.currentPaper      || 'Letter-L';
    currentBorderType = state.currentBorderType || 'dots';
    hideQRCard        = !!state.hideQRCard;
    showCardImages    = !!state.showCardImages;
    var hideQREl = document.getElementById('hide-qr-checkbox');
    if (hideQREl) hideQREl.checked = hideQRCard;
    var showImgEl = document.getElementById('show-card-images-checkbox');
    if (showImgEl) showImgEl.checked = showCardImages;
    cardIdCtr         = state.cardIdCtr         || cards.length;
    var paperSel  = document.getElementById('paper-select');
    var borderSel = document.getElementById('border-select');
    if (paperSel)  paperSel.value  = currentPaper;
    if (borderSel) borderSel.value = currentBorderType;
    return true;
  } catch(e) { return false; }
}

// ── FORMS ────────────────────────────────────────────

function buildPickerOptions(selectedKey, filter, cardSet, cardId) {
  var allCards = SET_CARD_DATA[cardSet] || BASE_SET_CARDS;
  var setCards = cardsForPokemon(cardSet);
  var f = filter.toLowerCase();
  var items = setCards.reduce(function(acc, bc) {
    var i = allCards.indexOf(bc);
    var label = bc.num + ' \u2014 ' + bc.name;
    if (!f || label.toLowerCase().indexOf(f) !== -1 || bc.name.toLowerCase().indexOf(f) !== -1) {
      acc.push({ i: i, label: label });
    }
    return acc;
  }, []);
  return buildPickerHTML(items,
    function(it) { return it.i === selectedKey; },
    function(it) { return it.label; },
    function(it) { return 'pickCard(event,' + cardId + ',' + it.i + ')'; },
    function(it) { return it.i; },
    'No cards found'
  );
}

var activePickerId = null;

function openCardPickerMousedown(e, cardId) {
  e.preventDefault();
  if (activePickerId === cardId) {
    var input = document.getElementById('picker-input-' + cardId);
    if (isMobileView() && input && input.hasAttribute('readonly')) { enablePickerTyping(input); return; }
    closeCardPicker(cardId); return;
  }
  if (activePickerId !== null) closeCardPicker(activePickerId);
  openCardPicker(cardId);
  if (!isMobileView()) enablePickerTyping(document.getElementById('picker-input-' + cardId));
}

function openSetPickerMousedown(e, cardId) {
  e.preventDefault();
  if (activeSetPickerId === cardId) {
    var input = document.getElementById('set-picker-input-' + cardId);
    if (isMobileView() && input && input.hasAttribute('readonly')) { enablePickerTyping(input); return; }
    closeSetPicker(cardId); return;
  }
  if (activeSetPickerId !== null) closeSetPicker(activeSetPickerId);
  openSetPicker(cardId);
  if (!isMobileView()) enablePickerTyping(document.getElementById('set-picker-input-' + cardId));
}

function openCardPicker(cardId) {
  var c = findCard(cardId);
  activePickerId = cardId;
  openPickerEl(
    document.getElementById('picker-drop-' + cardId),
    document.getElementById('picker-input-' + cardId),
    buildPickerOptions(c ? c.cardKey : -1, '', c ? c.set : '', cardId)
  );
}

function closeCardPicker(cardId) {
  if (activePickerId === cardId) activePickerId = null;
  var c = findCard(cardId);
  var scd = c && SET_CARD_DATA[c.set];
  closePickerEl(
    document.getElementById('picker-drop-' + cardId),
    document.getElementById('picker-input-' + cardId),
    c && c.cardKey >= 0 && scd && scd[c.cardKey]
      ? scd[c.cardKey].num + ' — ' + scd[c.cardKey].name
      : ''
  );
}

function filterCardPicker(cardId, value) {
  var drop = document.getElementById('picker-drop-' + cardId);
  var c = findCard(cardId);
  if (!drop) return;
  drop.innerHTML = buildPickerOptions(c ? c.cardKey : -1, value, c ? c.set : '', cardId);
  drop.classList.add('open');
}

function pickCard(e, cardId, idx) {
  e.preventDefault();
  var c = findCard(cardId);
  if (!c) return;
  if (c.cardKey === idx) { closeCardPicker(cardId); return; }
  var scd = SET_CARD_DATA[c.set] || [];
  var bc = scd[idx];
  if (!bc) return;
  c.cardKey  = idx;
  c.name     = bc.name;
  c.num      = bc.num || '';
  c.variants = [];
  c.specialVariants = [];
  closeCardPicker(cardId);
  render();
}

var activeSetPickerId = null;

function buildSetOptions(selectedSet, filter, cardId) {
  var f = filter.toLowerCase();
  var items = setsForPokemon().filter(function(s) { return !f || s.toLowerCase().indexOf(f) !== -1; });
  return buildPickerHTML(items,
    function(s) { return s === selectedSet; },
    function(s) { return s; },
    function(s) { return 'event.preventDefault();pickSet(event,' + cardId + ',\'' + s.replace(/'/g, "\\'") + '\')'; },
    function(s) { return s; },
    'No sets found'
  );
}

function openSetPicker(cardId) {
  var c = findCard(cardId);
  activeSetPickerId = cardId;
  openPickerEl(
    document.getElementById('set-picker-drop-' + cardId),
    document.getElementById('set-picker-input-' + cardId),
    buildSetOptions(c ? c.set : '', '', cardId)
  );
}

function closeSetPicker(cardId) {
  if (activeSetPickerId === cardId) activeSetPickerId = null;
  var c = findCard(cardId);
  closePickerEl(
    document.getElementById('set-picker-drop-' + cardId),
    document.getElementById('set-picker-input-' + cardId),
    c ? c.set || '' : ''
  );
}

function filterSetPicker(cardId, value) {
  var drop = document.getElementById('set-picker-drop-' + cardId);
  var c = findCard(cardId);
  if (!drop) return;
  drop.innerHTML = buildSetOptions(c ? c.set : '', value, cardId);
  drop.classList.add('open');
}

function pickSet(e, cardId, setName) {
  e.preventDefault();
  var c = findCard(cardId);
  if (!c) return;
  if (c.set === setName) { closeSetPicker(cardId); return; }
  c.set     = setName;
  c.cardKey = -1;
  c.name    = '';
  c.variants = [];
  c.specialVariants = [];
  var availableCards = cardsForPokemon(setName);
  if (availableCards.length === 1) {
    var allCards = SET_CARD_DATA[setName] || [];
    var bc = availableCards[0];
    c.cardKey = allCards.indexOf(bc);
    c.name    = bc.name;
    c.num     = bc.num || '';
  }
  closeSetPicker(cardId);
  render();
}

function renderForms() {
  var container = document.getElementById('card-forms');
  var clearBtn  = document.getElementById('clear-btn');
  var pdfBtn    = document.getElementById('pdf-btn');
  var addBtn    = document.getElementById('add-card-btn');

  clearBtn.style.display = '';
  if (addBtn) addBtn.style.display = '';
  var btnRow = document.getElementById('btn-row');
  if (btnRow) btnRow.style.display = pokemonFilter ? 'flex' : 'none';
  var hasAny = cards.some(function(c) { return c.name.trim(); });
  if (pdfBtn) pdfBtn.disabled = !hasAny;

  var pf = document.getElementById('pokemon-filter-section');
  if (pf) {
    pf.innerHTML = '<div id="poke-filter-wrap" style="margin-bottom:8px;">'
      + (miltankPlaying
        ? '<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;background:var(--surface2);border:1px solid var(--accent);border-radius:7px;padding:7px 10px;">'
          + '<input type="range" id="miltank-volume" min="0" max="100" value="' + miltankVolume + '"'
          + ' oninput="setMiltankVolume(this.value)"'
          + ' style="flex:1;accent-color:var(--accent);cursor:pointer;"/>'
          + '<button onclick="stopMiltankPlayer()" style="background:none;border:none;color:var(--muted);cursor:pointer;font-size:.75rem;padding:2px 5px;border-radius:4px;line-height:1;transition:color .15s;" onmouseover="this.style.color=\'var(--text)\'" onmouseout="this.style.color=\'var(--muted)\'">✕</button>'
          + '</div>'
        : '')
      + '<div id="poke-picker-row" style="display:flex;gap:6px;align-items:stretch;">'
      + '<div class="card-picker" id="poke-picker-wrap" style="flex:2;">'
      + '<input type="text" class="card-picker-input"'
      + ' id="poke-picker-input-global"'
      + ' placeholder="Search or select a Pokémon"'
      + ' value="' + esc(pokemonDisplayValue()) + '"'
      + ' autocomplete="off" readonly'
      + ' onmousedown="openPokemonPickerMousedown(event)"'
      + ' oninput="filterPokemonPicker(this.value)"'
      + ' onkeydown="pokemonPickerKeydown(event)"/>'
      + '<div class="card-picker-dropdown" id="poke-picker-drop-global">'
      + buildPokemonOptions('')
      + '</div></div>'
      + (pokemonFilter
        ? '<button class="btn btn-ghost" style="font-size:.65rem;gap:5px;flex-shrink:0;width:auto;padding:8px 10px;white-space:nowrap;" onmouseover="this.style.background=\'var(--accent)\';this.style.color=\'#111\';this.style.borderColor=\'var(--accent)\';" onmouseout="this.style.background=\'\';this.style.color=\'\';this.style.borderColor=\'\';" onclick="confirmFillAllCards()">'
          + '<svg width="12" height="12" viewBox="0 0 14 14" fill="none"><path d="M2 7h10M7 2l5 5-5 5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>'
          + 'Auto-Fill All Cards'
          + '</button>'
        : '')
      + '</div>'
      + '<div id="poke-checkboxes-row" style="display:flex;gap:12px;margin-top:12px;justify-content:center;">'
      + '<label style="display:flex;align-items:center;gap:6px;font-size:.7rem;color:var(--muted);cursor:pointer;">'
      + '<input type="checkbox" id="cameo-checkbox"'
      + (includeCameo ? ' checked' : '')
      + ' onchange="toggleCameo(this.checked)" style="width:auto;"/>'
      + 'Include Cameo Cards</label>'
      + '<label style="display:flex;align-items:center;gap:6px;font-size:.7rem;color:var(--muted);cursor:pointer;">'
      + '<input type="checkbox" id="serials-checkbox"'
      + (includeSerials ? ' checked' : '')
      + ' onchange="toggleSerials(this.checked)" style="width:auto;"/>'
      + 'Include EX Serial Numbers</label>'
      + '</div>'
      + '</div>';
  }

  container.innerHTML = cards.map(function(card, idx) {
    var imgSrc   = card.set ? getSymbolUrl(card.set) : '';
    var scd      = SET_CARD_DATA[card.set];
    var scCard   = scd && card.cardKey >= 0 ? scd[card.cardKey] : null;
    var noCard   = !!scd && !scCard;

    var variantOptions = [];
    if (scCard) {
      var cardData = scCard;
      if (Array.isArray(cardData.variants) && cardData.variants.length) {
        variantOptions = cardData.variants.slice();
      }
      variantOptions = applySerialSubstitution(variantOptions, cardData);
    }
    var specialVarOptions = scCard && Array.isArray(scCard.specialVariants) ? scCard.specialVariants : [];
    var variantHTML = '';
    if (variantOptions.length || specialVarOptions.length) {
      variantHTML = '<div class="ctrl" style="margin-top:10px;"><div class="edition-pills">'
        + variantOptions.map(function(v) {
            return '<button class="edition-pill' + ((card.variants || []).indexOf(v) !== -1 ? ' selected' : '')
              + '" data-variant="' + esc(v) + '" onclick="toggleVariant(this,this.dataset.variant,false)">'
              + esc(v) + '</button>';
          }).join('')
        + specialVarOptions.map(function(v) {
            return '<button class="edition-pill' + ((card.specialVariants || []).indexOf(v) !== -1 ? ' selected' : '')
              + '" data-variant="' + esc(v) + '" onclick="toggleVariant(this,this.dataset.variant,true)">'
              + esc(v) + '</button>';
          }).join('')
        + '</div></div>';
    }

    var cardHeaderSub = (function() {
      if (!card.name) return '';
      var fullNum = scCard ? scCard.num : '';
      return '<span style="color:var(--muted);margin-left:6px;">' + esc(card.name) + '</span>'
        + (fullNum ? '<span style="color:var(--muted);margin-left:4px;font-size:.85em;"> - ' + esc(fullNum) + '</span>' : '');
    }());
    var cardInputVal = scCard ? ' value=' + JSON.stringify(scCard.num + ' — ' + scCard.name) : ' value=""';

    return '<div class="card-form" id="form-' + card.id + '" data-cardid="' + card.id + '">'
      + '<div class="card-form-header" onclick="toggleCardCollapse(' + card.id + ',event)">'
      +   '<span style="display:flex;align-items:center;gap:4px;">'
      +     '<button class="collapse-btn" tabindex="-1">&#9660;</button>'
      +     '<span style="color:var(--accent);">Card #' + (idx + 1) + '</span>'
      +     cardHeaderSub
      +   '</span>'
      +   '<span style="display:flex;align-items:center;gap:2px;">'
      +   '<button class="remove-btn" onclick="duplicateCard(' + card.id + ')" title="Duplicate" style="color:var(--muted);">⧉</button>'
      +   '<button class="remove-btn" onclick="removeCard(' + card.id + ')">&#x2715;</button>'
      +   '</span>'
      + '</div>'
      + '<div class="card-form-body">'
      + '<div class="dropdowns-row">'
      + '<div class="ctrl">'
      +   '<div class="set-row">'
      +     '<div class="card-picker" id="set-picker-' + card.id + '">'
      +       '<input type="text" class="card-picker-input"'
      +         ' id="set-picker-input-' + card.id + '"'
      +         ' placeholder="Search or select a set..."'
      +         ' value="' + esc(card.set || '') + '"'
      +         ' autocomplete="off" readonly'
      +         ' onmousedown="openSetPickerMousedown(event,' + card.id + ')"'
      +         ' oninput="filterSetPicker(' + card.id + ',this.value)"'
      +         ' onkeydown="setPickerKeydown(event,' + card.id + ')"/>'
      +       '<div class="card-picker-dropdown" id="set-picker-drop-' + card.id + '">'
      +       '</div>'
      +     '</div>'
      +     '<img class="set-sym-thumb" src="' + imgSrc + '" style="' + (imgSrc ? '' : 'display:none;') + '" alt=""/>'
      +   '</div>'
      + '</div>'
      + (SET_CARD_DATA[card.set]
        ? '<div class="ctrl" style="margin-top:10px;">'
          + '<div class="card-picker" id="picker-' + card.id + '">'
          + '<input type="text" class="card-picker-input"'
          + ' id="picker-input-' + card.id + '"'
          + ' placeholder="Search or select a card..."'
          + cardInputVal
          + ' autocomplete="off" readonly'
          + ' onmousedown="openCardPickerMousedown(event, ' + card.id + ')"'
          + ' oninput="filterCardPicker(' + card.id + ', this.value)"'
          + ' onkeydown="cardPickerKeydown(event, ' + card.id + ')"/>'
          + '<div class="card-picker-dropdown" id="picker-drop-' + card.id + '">'
          + '</div></div></div>'
        : '<div style="font-size:.62rem;color:var(--muted);padding:6px 2px;">Select a set to choose a card.</div>')
      + '</div>'
      + (!noCard ? variantHTML : '')
      + (function() {
          if (!scCard) return '';
          var hasVariants = variantOptions.length > 0;
          var hasSpecial  = specialVarOptions.length > 0;
          if (!hasVariants && !hasSpecial) return '';
          var noneSelected = (card.variants || []).length === 0 && (card.specialVariants || []).length === 0;
          return noneSelected
            ? '<div class="variant-warning" style="font-size:.6rem;color:#c47a82;margin-top:8px;text-align:center;">⚠ No Variants Selected</div>'
            : '';
        }())
      + '</div>'
      + '</div></div>';
  }).join('');

  Object.keys(collapsedIds).forEach(function(id) {
    var el = document.getElementById('form-' + id);
    if (el) el.classList.add('collapsed');
  });
}

// ── PREVIEW ──────────────────────────────────────────

function getTotalVariantCount(card) {
  var sclR = SET_CARD_DATA[card.set];
  if (!sclR || card.cardKey < 0 || !sclR[card.cardKey]) return 0;
  var cdR = sclR[card.cardKey];
  var stdVars = applySerialSubstitution(Array.isArray(cdR.variants) ? cdR.variants.slice() : [], cdR);
  var specVars = Array.isArray(cdR.specialVariants) ? cdR.specialVariants : [];
  return stdVars.length + specVars.length;
}

var _lastScale = 1;
var _lastScaledW = 0;
var _lastScaledH = 0;

function scalePages() {
  var preview = document.getElementById('preview');
  if (!preview) return;
  var wraps = preview.querySelectorAll('.page-scale-wrap');
  if (!wraps.length) return;
  var rect = preview.getBoundingClientRect();
  var availW = rect.width - 56;
  var availH = rect.height - 80;
  if (availW <= 0) return;
  wraps.forEach(function(wrap) {
    var pageW = parseInt(wrap.getAttribute('data-page-w'), 10);
    var pageH = parseInt(wrap.getAttribute('data-page-h'), 10);
    if (!pageW || !pageH) return;
    var scaleW = availW / pageW;
    var scaleH = availH > 0 ? availH / pageH : 1;
    var scale = Math.min(1, scaleW, scaleH);
    var pageWrap = wrap.querySelector('.page-wrap');
    if (!pageWrap) return;
    pageWrap.style.transform = 'scale(' + scale + ')';
    pageWrap.style.transformOrigin = 'top left';
    var scaledW = Math.round(pageW * scale);
    var scaledH = Math.round(pageH * scale);
    wrap.style.width  = scaledW + 'px';
    wrap.style.height = scaledH + 'px';
    _lastScale   = scale;
    _lastScaledW = scaledW;
    _lastScaledH = scaledH;
  });
}

function applyLastScale() {
  if (!_lastScaledH) return;
  var wraps = document.querySelectorAll('.page-scale-wrap');
  wraps.forEach(function(wrap) {
    var pageWrap = wrap.querySelector('.page-wrap');
    if (!pageWrap) return;
    pageWrap.style.transform = 'scale(' + _lastScale + ')';
    pageWrap.style.transformOrigin = 'top left';
    wrap.style.width  = _lastScaledW + 'px';
    wrap.style.height = _lastScaledH + 'px';
  });
}
var _resizeTimer = null;
window.addEventListener('resize', function() {
  if (_resizeTimer) clearTimeout(_resizeTimer);
  _resizeTimer = setTimeout(scalePages, 80);
});

// Returns the last null index in the chunk (for promo slot placement), or -1.
function findPromoSlotIndex(chunk) {
  if (!chunk || chunk[chunk.length - 1] !== null) return -1;
  for (var li = chunk.length - 1; li >= 0; li--) {
    if (chunk[li] === null) return li;
  }
  return -1;
}

function renderPreview() {
  var container  = document.getElementById('pages-container');

  var p        = PAPER_DATA[currentPaper];
  var perPage  = p.cols * p.rows;
  var pageWpx  = Math.round(p.pageW  * DPI);
  var pageHpx  = Math.round(p.pageH  * DPI);
  var marginPx = Math.round(p.margin * DPI);

  var slots = buildSlots();

  if (slots.length === 0) {
    var emptySlots = '';
    for (var ei = 0; ei < p.cols * p.rows; ei++) emptySlots += '<div class="card-slot empty"><span class="corner-dot tl"></span><span class="corner-dot tr"></span><span class="corner-dot bl"></span><span class="corner-dot br"></span></div>';
    container.innerHTML =
      '<div class="page-scale-wrap" data-page-w="' + pageWpx + '" data-page-h="' + pageHpx + '">'
      + '<div class="page-wrap" style="width:' + pageWpx + 'px;height:' + pageHpx + 'px;padding:' + marginPx + 'px;">'
      + '<div class="card-grid border-' + currentBorderType + '" style="'
      + 'grid-template-columns:repeat(' + p.cols + ',' + CARD_W_PX + 'px);'
      + 'grid-template-rows:repeat(' + p.rows + ',' + CARD_H_PX + 'px);'
      + 'gap:0;">'
      + emptySlots + '</div></div></div>';
    document.getElementById('pg-label').textContent = '0 cards · 0 pages';
    applyLastScale();
    requestAnimationFrame(scalePages);
    return;
  }

  var chunks = [];
  for (var i = 0; i < slots.length; i += perPage) {
    var chunk = slots.slice(i, i + perPage);
    while (chunk.length < perPage) chunk.push(null);
    chunks.push(chunk);
  }

  var lastChunk = chunks[chunks.length - 1];
  var promoSlotIndex = findPromoSlotIndex(lastChunk);

  container.innerHTML = chunks.map(function(chunk, pi) {
    var isLastPage = pi === chunks.length - 1;
    return '<div class="page-label">Page ' + (pi + 1) + ' of ' + chunks.length + '</div>'
      + '<div class="page-scale-wrap" data-page-w="' + pageWpx + '" data-page-h="' + pageHpx + '">'
      + '<div class="page-wrap" style="'
      +   'width:'  + pageWpx  + 'px;'
      +   'height:' + pageHpx  + 'px;'
      +   'padding:' + marginPx + 'px;'
      + '">'
      +   '<div class="card-grid border-' + currentBorderType + '" style="'
      +     'grid-template-columns:repeat(' + p.cols + ',' + CARD_W_PX + 'px);'
      +     'grid-template-rows:repeat('    + p.rows + ',' + CARD_H_PX + 'px);'
      +     'gap:0;'
      +   '">'
      +   chunk.map(function(slot, si) {
            if (!hideQRCard && isLastPage && promoSlotIndex >= 0 && si === promoSlotIndex) {
              return '<div class="card-slot promo">' + CORNER_DOTS
                + '<div class="promo-inner">'
                + '<div class="promo-title">Got Miltank?</div>'
                + '<canvas class="promo-logo" id="promo-canvas-logo"></canvas>'
                + '<canvas class="promo-qr" id="promo-canvas-qr"></canvas>'
                + '</div></div>';
            }
            return buildSlotHTML(slot ? slot.card : null, slot ? slot.variant : null, slot ? slot.specialVariant : null);
          }).join('')
      +   '</div>'
      + '</div></div>';
  }).join('');

  var n = slots.length;
  document.getElementById('pg-label').textContent =
    n + ' card' + (n !== 1 ? 's' : '') + ' \xb7 ' + chunks.length + ' page' + (chunks.length !== 1 ? 's' : '');
  applyLastScale();
  requestAnimationFrame(function() { scalePages(); fitCardNames(); });
  if (promoSlotIndex >= 0) renderPromoCanvases(document);
  saveState();
}

function renderPromoCanvases(root, scale) {
  var LOGO_URL = 'https://reiterei.github.io/got-miltank/bw-logo.png';
  var QR_URL   = 'https://reiterei.github.io/got-miltank/qr-code.png';
  var dpr = scale || window.devicePixelRatio || 1;

  function drawToCanvas(canvas, url, displayW, displayH) {
    return new Promise(function(resolve) {
      var img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = function() {
        var bw = Math.round(displayW * dpr);
        var bh = Math.round(displayH * dpr);
        canvas.width  = bw;
        canvas.height = bh;
        canvas.style.width  = displayW + 'px';
        canvas.style.height = displayH + 'px';
        var ctx = canvas.getContext('2d');
        ctx.imageSmoothingEnabled = false;
        ctx.drawImage(img, 0, 0, bw, bh);
        resolve();
      };
      img.onerror = resolve;
      img.src = url;
    });
  }

  var promises = [];
  var logoCanvases = root.querySelectorAll ? root.querySelectorAll('#promo-canvas-logo') : [];
  var qrCanvases   = root.querySelectorAll ? root.querySelectorAll('#promo-canvas-qr')   : [];
  logoCanvases.forEach(function(c) { promises.push(drawToCanvas(c, LOGO_URL, 80, 80)); });
  qrCanvases.forEach(function(c)   { promises.push(drawToCanvas(c, QR_URL,   60, 60)); });
  return Promise.all(promises);
}

function fitCardNames() {
  function shrink(sel, minPx) {
    document.querySelectorAll(sel).forEach(function(el) {
      el.style.fontSize = '';
      var maxW = el.parentElement ? el.parentElement.offsetWidth : el.offsetWidth;
      var fontSize = parseFloat(getComputedStyle(el).fontSize);
      while (el.scrollWidth > maxW + 1 && fontSize > minPx) {
        fontSize -= 0.5;
        el.style.fontSize = fontSize + 'px';
      }
    });
  }
  shrink('.cs-top .cs-name', 9);
  shrink('.cs-bottom .cs-set', 8);
}

// ── UTILS ────────────────────────────────────────────
var _escMap = { '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;' };
function esc(s) {
  return String(s || '').replace(/[&<>"]/g, function(c) { return _escMap[c]; });
}

function buildSlots() {
  var slots = [];
  for (var ci = 0; ci < cards.length; ci++) {
    var cd = cards[ci];
    var variantOpts = [];
    var hasVariantData = false;
    var sclP = SET_CARD_DATA[cd.set];
    var specialOpts = [];
    var cdData = sclP && cd.cardKey >= 0 ? sclP[cd.cardKey] : null;
    if (cdData) {
      if (Array.isArray(cdData.variants) && cdData.variants.length) {
        variantOpts = cdData.variants.slice();
        hasVariantData = true;
      }
      variantOpts = applySerialSubstitution(variantOpts, cdData);
      if (Array.isArray(cdData.specialVariants) && cdData.specialVariants.length) specialOpts = cdData.specialVariants;
    }
    var hasPrChoice = variantOpts.length > 1 && cd.variants && cd.variants.length > 0;
    var hasPrOnly1  = !hasVariantData && specialOpts.length === 0;
    var hasPr1Sel   = variantOpts.length === 1 && cd.variants && cd.variants.length > 0;
    if (hasPrOnly1 || hasPr1Sel || hasPrChoice) {
      var visibleSelected = variantOpts.length > 1 && cd.variants && cd.variants.length
        ? cd.variants.filter(function(v) { return variantOpts.indexOf(v) !== -1; })
        : variantOpts.length === 1 ? [variantOpts[0]] : [null];
      visibleSelected.forEach(function(variant) {
        slots.push({ card: cd, variant: variant, specialVariant: null });
      });
    }
    var selSpecials = specialOpts.length && cd.specialVariants && cd.specialVariants.length ? cd.specialVariants : [];
    selSpecials.forEach(function(sh) {
      slots.push({ card: cd, variant: null, specialVariant: sh });
    });
  }
  return slots;
}

function buildSlotHTML(card, variantOverride, specialVariantOverride) {
  if (!card || !card.name || !card.name.trim()) {
    return '<div class="card-slot empty">' + CORNER_DOTS + '</div>';
  }
  var sym       = card.set ? getSymbolUrl(card.set) : null;
  var showLabel = getTotalVariantCount(card) > 1;

  // Card image mode
  if (showCardImages) {
    var scd2 = SET_CARD_DATA[card.set];
    var scCard2 = scd2 && card.cardKey >= 0 ? scd2[card.cardKey] : null;
    var imgUrl = card.num ? getScrydexImageUrl(card.set, card.num, scCard2 ? scCard2.imageUrl : '') : null;
    if (imgUrl) {
      return '<div class="card-slot card-image-slot">' + CORNER_DOTS
        + '<img class="cs-card-image" src="' + imgUrl + '" alt="' + esc(card.name) + '"'
        + ' onerror="this.closest(\'.card-image-slot\').classList.add(\'img-failed\');this.style.display=\'none\';"'
        + '/>'
        + (showLabel && (variantOverride || specialVariantOverride)
            ? '<div class="cs-image-variant-label">' + esc(variantOverride || specialVariantOverride) + '</div>'
            : '')
        + '<div class="cs-inner cs-image-fallback">'
        +   '<div class="cs-top">'
        +     '<div class="cs-name">' + esc(card.name) + '</div>'
        +     (card.num ? '<div class="cs-number">' + esc(card.num) + '</div>' : '<div class="cs-number blank">\u2014</div>')
        +   '</div>'
        +   '<div class="cs-symbol-wrap"><img class="cs-symbol" src="' + (sym || BLANK_SYMBOL_URI) + '" alt=""/></div>'
        +   '<div class="cs-bottom">'
        +     (card.set ? '<div class="cs-set">' + esc(card.set) + '</div>' : '<div class="cs-set" style="color:#ccc;">\u2014</div>')
        +     (showLabel && variantOverride        ? '<div class="cs-variant-label">' + esc(variantOverride)        + '</div>' : '')
        +     (showLabel && specialVariantOverride ? '<div class="cs-variant-label">' + esc(specialVariantOverride) + '</div>' : '')
        +   '</div>'
        + '</div>'
        + '</div>';
    }
  }

  return '<div class="card-slot">' + CORNER_DOTS
    + '<div class="cs-inner">'
    +   '<div class="cs-top">'
    +     '<div class="cs-name">' + esc(card.name) + '</div>'
    +     (card.num
            ? '<div class="cs-number">' + esc(card.num) + '</div>'
            : '<div class="cs-number blank">\u2014</div>')
    +   '</div>'
    +   '<div class="cs-symbol-wrap">'
    +     '<img class="cs-symbol" src="' + (sym || BLANK_SYMBOL_URI) + '" alt=""/>'
    +   '</div>'
    +   '<div class="cs-bottom">'
    +     (card.set
            ? '<div class="cs-set">' + esc(card.set) + '</div>'
            : '<div class="cs-set" style="color:#ccc;">\u2014</div>')
    +     (showLabel && variantOverride        ? '<div class="cs-variant-label">' + esc(variantOverride)        + '</div>' : '')
    +     (showLabel && specialVariantOverride ? '<div class="cs-variant-label">' + esc(specialVariantOverride) + '</div>' : '')
    +   '</div>'
    + '</div></div>';
}

// ── SAVE PDF ─────────────────────────────────────────

function isMobileView() {
  var preview = document.getElementById('preview');
  return preview && getComputedStyle(preview).display === 'none';
}

async function savePDF() {
  var pdfBtn = document.getElementById('pdf-btn');
  if (pdfBtn) {
    pdfBtn.disabled = true;
    pdfBtn.innerHTML = '<svg width="13" height="13" viewBox="0 0 14 14" fill="none"><circle cx="7" cy="7" r="5" stroke="currentColor" stroke-width="1.2" stroke-dasharray="8 6" transform="rotate(-90 7 7)"><animateTransform attributeName="transform" type="rotate" from="0 7 7" to="360 7 7" dur="0.8s" repeatCount="indefinite"/></circle></svg> Generating\u2026';
  }

  var offscreen = null;
  if (isMobileView()) {
    var p2       = PAPER_DATA[currentPaper];
    var pageWpx2 = Math.round(p2.pageW  * DPI);
    var pageHpx2 = Math.round(p2.pageH  * DPI);
    var marginPx2 = Math.round(p2.margin * DPI);

    var slots2   = buildSlots();
    var perPage2 = p2.cols * p2.rows;
    var chunks2  = [];
    for (var i2 = 0; i2 < slots2.length; i2 += perPage2) {
      var chunk2 = slots2.slice(i2, i2 + perPage2);
      while (chunk2.length < perPage2) chunk2.push(null);
      chunks2.push(chunk2);
    }

    var lastChunk2 = chunks2[chunks2.length - 1];
    var promoSlotIndex2 = findPromoSlotIndex(lastChunk2);

    offscreen = document.createElement('div');
    offscreen.style.cssText = 'position:fixed;left:-99999px;top:0;width:' + pageWpx2 + 'px;opacity:0;pointer-events:none;';
    chunks2.forEach(function(chunk2, pi2) {
      var isLastPage2 = pi2 === chunks2.length - 1;
      var pageEl = document.createElement('div');
      pageEl.className = 'page-wrap';
      pageEl.style.cssText = 'width:' + pageWpx2 + 'px;height:' + pageHpx2 + 'px;padding:' + marginPx2 + 'px;background:white;display:flex;align-items:center;justify-content:center;';
      pageEl.innerHTML = '<div class="card-grid border-' + currentBorderType + '" style="'
        + 'grid-template-columns:repeat(' + p2.cols + ',' + CARD_W_PX + 'px);'
        + 'grid-template-rows:repeat(' + p2.rows + ',' + CARD_H_PX + 'px);'
        + 'gap:0;">'
        + chunk2.map(function(slot, si2) {
            if (!hideQRCard && isLastPage2 && promoSlotIndex2 >= 0 && si2 === promoSlotIndex2) {
              return '<div class="card-slot promo">' + CORNER_DOTS
                + '<div class="promo-inner">'
                + '<div class="promo-title">Got Miltank?</div>'
                + '<canvas class="promo-logo" id="promo-canvas-logo"></canvas>'
                + '<canvas class="promo-qr" id="promo-canvas-qr"></canvas>'
                + '</div></div>';
            }
            return buildSlotHTML(slot ? slot.card : null, slot ? slot.variant : null, slot ? slot.specialVariant : null);
          }).join('')
        + '</div>';
      offscreen.appendChild(pageEl);
    });
    document.body.appendChild(offscreen);
    if (promoSlotIndex2 >= 0) await renderPromoCanvases(offscreen, 3);
    await new Promise(function(resolve) { requestAnimationFrame(function() { requestAnimationFrame(resolve); }); });
  }

  try {
    var p       = PAPER_DATA[currentPaper];
    var pageWmm = p.pageW * 25.4;
    var pageHmm = p.pageH * 25.4;
    var orient  = p.pageW > p.pageH ? 'landscape' : 'portrait';

    var { jsPDF } = window.jspdf;
    var doc = new jsPDF({ orientation: orient, unit: 'mm', format: [pageWmm, pageHmm], compress: true });

    var pageWraps = offscreen
      ? offscreen.querySelectorAll('.page-wrap')
      : document.querySelectorAll('#pages-container .page-wrap');
    if (!pageWraps.length) return;

    var pdfRoot = offscreen || document.getElementById('pages-container');
    if (pdfRoot && pdfRoot.querySelector('#promo-canvas-logo')) {
      await renderPromoCanvases(pdfRoot, 3);
    }

    for (var i = 0; i < pageWraps.length; i++) {
      if (i > 0) doc.addPage([pageWmm, pageHmm], orient);

      var wrap = pageWraps[i];
      var scaleWrap = offscreen ? null : wrap.parentElement;
      var prevTransform = wrap.style.transform;
      var prevWrapW = scaleWrap ? scaleWrap.style.width : null;
      var prevWrapH = scaleWrap ? scaleWrap.style.height : null;
      var prevOverflow = scaleWrap ? scaleWrap.style.overflow : null;
      wrap.style.transform = 'none';
      if (scaleWrap) {
        scaleWrap.style.width = wrap.style.width;
        scaleWrap.style.height = wrap.style.height;
        scaleWrap.style.overflow = 'visible';
      }

      var canvas = await html2canvas(wrap, {
        scale: 3,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false
      });

      wrap.style.transform = prevTransform;
      if (scaleWrap) {
        scaleWrap.style.width    = prevWrapW;
        scaleWrap.style.height   = prevWrapH;
        scaleWrap.style.overflow = prevOverflow;
      }

      var imgData = canvas.toDataURL('image/jpeg', 0.95);
      doc.addImage(imgData, 'JPEG', 0, 0, pageWmm, pageHmm);
    }

    var pokeSlug = '';
    if (pokemonFilter) {
      var dexIdx = parseInt(pokemonFilter.replace('#',''), 10) - 1;
      var pokeName = (dexIdx >= 0 && POKEMON_LIST[dexIdx]) ? POKEMON_LIST[dexIdx] : pokemonFilter;
      pokeSlug = pokeName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase() + '-';
    }
    doc.save(pokeSlug + 'card-templates.pdf');
  } catch (err) {
    alert('PDF generation failed: ' + err.message);
    console.error(err);
  } finally {
    if (offscreen && offscreen.parentNode) offscreen.parentNode.removeChild(offscreen);
    var previewRoot = document.getElementById('pages-container');
    if (previewRoot && previewRoot.querySelector('#promo-canvas-logo')) {
      renderPromoCanvases(previewRoot);
    }
    if (pdfBtn) {
      var hasAny = cards.some(function(c) { return c.name && c.name.trim(); });
      pdfBtn.disabled = !hasAny;
      pdfBtn.innerHTML = '<svg width="13" height="13" viewBox="0 0 14 14" fill="none">'
        + '<path d="M2 2.5A1.5 1.5 0 013.5 1h5l3 3v8a1.5 1.5 0 01-1.5 1.5h-7A1.5 1.5 0 012 11.5V2.5z" stroke="currentColor" stroke-width="1.2"/>'
        + '<path d="M8.5 1v3h3M5 7.5h4M5 9.5h2.5" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/>'
        + '</svg> Save PDF';
    }
  }
}

// ── INIT ─────────────────────────────────────────────

function parseCSV(text) {
  var lines = text.split('\n');
  var headers = splitCSVRow(lines[0]);
  var rows = [];
  for (var i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    var vals = splitCSVRow(lines[i]);
    var row = {};
    headers.forEach(function(h, j) { row[h.trim()] = (vals[j] || '').trim(); });
    rows.push(row);
  }
  return rows;
}

function splitCSVRow(row) {
  var result = [];
  var inQuote = false;
  var cur = '';
  for (var i = 0; i < row.length; i++) {
    var c = row[i];
    if (c === '"') { inQuote = !inQuote; }
    else if (c === ',' && !inQuote) { result.push(cur); cur = ''; }
    else { cur += c; }
  }
  result.push(cur);
  return result;
}

function buildDataFromSheets(cardRows, setRows, pokemonRows) {
  function splitCol(val) {
    return val ? val.split(',').map(function(v) { return v.trim(); }).filter(Boolean) : [];
  }

  SET_DATA = {};
  setRows.forEach(function(row) {
    var name = row['Set Name'];
    if (!name) return;
    SET_DATA[name] = { symbol: (row['Symbol'] || '').trim(), scrydexId: (row['Scrydex ID'] || '').trim() };
  });

  SET_CARD_DATA = {};
  var setsOrder = [];
  cardRows.forEach(function(row) {
    var setName = row['Set Name'];
    if (!setName) return;
    if (!SET_CARD_DATA[setName]) { SET_CARD_DATA[setName] = []; setsOrder.push(setName); }
    SET_CARD_DATA[setName].push({
      num:            row['Card Number']          || '',
      name:           row['Card Name']            || '',
      variants:       splitCol(row['Standard Variants']),
      specialVariants:splitCol(row['Special Variants']),
      serialVariants: splitCol(row['Serial Number Variants']),
      primary:        row['Primary Pokemon']      || '',
      cameo:          row['Cameo Pokemon']        || '',
      imageUrl:       (row['Image URL'] || '').trim(),
    });
  });
  KNOWN_SETS = setsOrder;

  if (pokemonRows && pokemonRows.length) {
    pokemonRows.sort(function(a, b) {
      return parseInt(a['Dex Number'] || 0, 10) - parseInt(b['Dex Number'] || 0, 10);
    });
    POKEMON_LIST = pokemonRows.map(function(row) { return (row['Name'] || '').trim(); }).filter(Boolean);
  } else {
    var pokeSet = {};
    cardRows.forEach(function(row) {
      splitCol(row['Primary Pokemon']).forEach(function(p) { pokeSet[p] = 1; });
      splitCol(row['Cameo Pokemon']).forEach(function(p)   { pokeSet[p] = 1; });
    });
    POKEMON_LIST = Object.keys(pokeSet).sort();
  }

  _dexLabels = POKEMON_LIST.map(function(name, i) {
    return '#' + String(i + 1).padStart(4, '0') + ' - ' + name;
  });

  BASE_SET_CARDS = SET_CARD_DATA['Base Set'] || [];
}

function showLoadingError(msg) {
  var el = document.getElementById('loading-screen');
  if (el) el.innerHTML = '<div style="text-align:center;padding:40px;color:#ff7070;">'
    + '<div style="font-size:1.2rem;margin-bottom:12px;">⚠ Failed to load data</div>'
    + '<div style="font-size:.75rem;color:var(--muted);max-width:400px;margin:0 auto;">' + msg + '</div>'
    + '<div style="margin-top:20px;font-size:.7rem;color:var(--muted2);">Check that your Google Sheets URLs are correct in the CONFIG section of the HTML file.</div>'
    + '<div style="margin-top:24px;display:flex;gap:10px;justify-content:center;">'
    + '<button onclick="location.reload()" style="padding:8px 16px;border-radius:6px;border:1px solid var(--border);background:transparent;color:var(--text);font-family:inherit;font-size:.7rem;cursor:pointer;">↺ Retry</button>'
    + '<button onclick="loadFallbackData()" style="padding:8px 16px;border-radius:6px;border:1px solid var(--border);background:transparent;color:var(--muted);font-family:inherit;font-size:.7rem;cursor:pointer;">Use Fallback Data</button>'
    + '</div>'
    + '</div>';
}

function hideLoadingScreen() {
  var el = document.getElementById('loading-screen');
  if (el) el.style.display = 'none';
  var app = document.querySelector('.app');
  if (app) app.style.display = '';
}

async function loadData() {
  if (!CONFIG.CARDS_SHEET_URL || CONFIG.CARDS_SHEET_URL === 'YOUR_CARDS_SHEET_CSV_URL_HERE') {
    console.warn('Google Sheets URLs not configured — using empty data.');
    initApp();
    return;
  }

  var timeoutId = setTimeout(function() {
    console.warn('Sheet load timed out — using fallback data.');
    showLoadingError('Loading timed out. Check your Google Sheets URLs and make sure the sheets are published.');
  }, 10000);

  try {
    var [cardsResp, setsResp, pokemonResp] = await Promise.all([
      fetch(CONFIG.CARDS_SHEET_URL),
      fetch(CONFIG.SETS_SHEET_URL),
      fetch(CONFIG.POKEMON_SHEET_URL),
    ]);

    if (!cardsResp.ok)   throw new Error('Cards sheet returned '   + cardsResp.status);
    if (!setsResp.ok)    throw new Error('Sets sheet returned '    + setsResp.status);
    if (!pokemonResp.ok) throw new Error('Pokémon sheet returned ' + pokemonResp.status);

    var cardsText   = await cardsResp.text();
    var setsText    = await setsResp.text();
    var pokemonText = await pokemonResp.text();

    var cardRows   = parseCSV(cardsText);
    var setRows    = parseCSV(setsText);
    var pokemonRows = parseCSV(pokemonText);

    buildDataFromSheets(cardRows, setRows, pokemonRows);
    clearTimeout(timeoutId);
    initApp();

  } catch(err) {
    clearTimeout(timeoutId);
    console.error('Failed to load sheet data:', err);
    showLoadingError('Failed to load data: ' + err.message);
  }
}

function loadFallbackData() {
  SET_DATA = {
    'MissingSet': {
      symbol:   'https://reiterei.github.io/got-miltank/symbols/MissingSet.PNG',
    }
  };
  SET_CARD_DATA = {
    'MissingSet': [
      { num: '0/0', name: 'MissingNo.', variants: ['Non-Holo'], specialVariants: ['Cosmos Holo'], serialVariants: ['000-000-001', '000-000-002'], primary: '#0001', cameo: '' }
    ]
  };
  KNOWN_SETS    = ['MissingSet'];
  POKEMON_LIST  = ['MissingNo.'];
  _dexLabels    = ['#0001 - MissingNo.'];
  BASE_SET_CARDS = SET_CARD_DATA['MissingSet'];
  initApp();
}

loadData();
