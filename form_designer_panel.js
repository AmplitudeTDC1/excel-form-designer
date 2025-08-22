/* form_designer_panel.js
   Unified (V1 base + V2 per-element span + Excel binding + hidden sheet autosave)

   Notes:
   - Expects (optionally) bindingManager.js loaded before this file (window.BindingManager).
   - Safe to run without Office.js; Excel persistence is no-op if not available.
*/

/* ---------- Utility ---------- */
const $  = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));
const on = (el, ev, fn, opts) => el && el.addEventListener(ev, fn, opts);
const clamp = (v, min, max) => Math.max(min, Math.min(max, Number.isFinite(v) ? v : 0));

/* Debounce for autosave */
function debounce(fn, delay) {
  let t;
  return function(...args){
    clearTimeout(t);
    t = setTimeout(()=>fn.apply(this,args), delay);
  };
}

/* Save status helper (optional #saveStatus element) */
function showSaveStatus(msg){
  const el = $('#saveStatus');
  if (el) el.textContent = msg;
}

/* ---------- BindingManager (optional) ---------- */
const BindingShim = {
  isReady: () => false,
  async write(_addr, _val) { /* no-op */ },
  async read(_addr) { return ''; },
  async ensureBinding(addr) { return addr || ''; },
  subscribe(_id, _cb) { /* no-op */ },
  unsubscribe(_id) { /* no-op */ }
};
const BM = (window.BindingManager && typeof window.BindingManager.write === 'function')
  ? window.BindingManager
  : BindingShim;

/* ---------- Elements ---------- */
const toolbox      = $$('.tool');
const canvas       = $('#canvas');
const formRoot     = $('#formRoot');
const gridOverlay  = $('#grid');

const layoutMode   = $('#layoutMode'); // 'flow' or 'free'
const columnsSelect= $('#columns');
const gridToggle   = $('#gridToggle');
const snapToggle   = $('#snapToggle');
const gridSizeInput= $('#gridSize');

const selectedInfo    = $('#selectedInfo');
const elementSettings = $('#elementSettings');
const globalSettings  = $('#globalSettings');

/* Element styling controls */
const elLabel      = $('#elLabel');
const elPlaceholder= $('#elPlaceholder');
const elRequired   = $('#elRequired');
const elWidth      = $('#elWidth');
const elHeight     = $('#elHeight');
const elFont       = $('#elFont');
const elFontSize   = $('#elFontSize');
const elLabelSize  = $('#elLabelSize');
const elTextColor  = $('#elTextColor');
const elBgColor    = $('#elBgColor');
const elBorderColor= $('#elBorderColor');
const elRadius     = $('#elRadius');
const elPadding    = $('#elPadding');

/* Per-element Column Span control (injected if absent) */
let elColSpan = $('#elColSpan');

/* Binding UI (optional) */
const elBindAddress = $('#elBindAddress');
const bindReadBtn   = $('#bindReadBtn');
const bindWriteBtn  = $('#bindWriteBtn');
const elBindListen  = $('#elBindListen');

/* Options editor */
const optionsEditor   = $('#optionsEditor');
const optionsList     = $('#optionsList');
const newOptionText   = $('#newOptionText');
const addOptionBtn    = $('#addOptionBtn');

/* Global settings */
const formBgColor   = $('#formBgColor');
const formBgImage   = $('#formBgImage');
const formWidth     = $('#formWidth');
const formPadding   = $('#formPadding');

/* Export / import / save */
const exportJsonBtn   = $('#exportJson');
const importJsonBtn   = $('#importJson');
const importFileInput = $('#importFileInput');
const saveLayoutBtn   = $('#saveLayout');
const clearCanvasBtn  = $('#clearCanvas');

/* Optional tabs */
const tabGlobal  = $('#tabGlobal');
const tabElement = $('#tabElement');

/* ---------- State ---------- */
let selectedEl  = null;
let dragDataType= null;
let isFreeMode  = false;
let gridSize    = parseInt(gridSizeInput?.value, 10) || 12;
let idCounter   = 0;

/* ---------- Helpers ---------- */
function getCurrentColumns(){
  return Math.max(1, parseInt(columnsSelect?.value,10) || 1);
}

/* Auto-inject Column Span control if missing */
(function ensureColSpanUI(){
  if (!elementSettings) return;
  if (!elColSpan) {
    const wrapper = document.createElement('label');
    wrapper.style.display = 'block';
    wrapper.style.marginTop = '8px';
    wrapper.innerHTML = `
      <span style="display:block; margin-bottom:4px;">Column Span</span>
      <input type="number" id="elColSpan" min="1" value="1" />
    `;
    const twoCols = elementSettings.querySelector('.two-cols');
    if (twoCols) twoCols.insertAdjacentElement('afterend', wrapper);
    else elementSettings.insertAdjacentElement('afterbegin', wrapper);
    elColSpan = $('#elColSpan');
  }
})();

function genId(prefix='fld'){ idCounter++; return `${prefix}_${Date.now().toString(36)}_${idCounter}`; }

/* ---------- Field factory ---------- */
function createField(type, data = {}) {
  const wrapper = document.createElement('div');
  wrapper.className = 'form-field';
  wrapper.setAttribute('data-type', type);
  wrapper.setAttribute('data-id', data.id || genId(type));

  const meta = Object.assign({
    id: wrapper.getAttribute('data-id'),
    label: (type === 'heading') ? 'Heading' : (type === 'button' ? 'Button' : 'Label'),
    placeholder: '',
    required: false,
    widthPct: 100,
    height: 0,
    font: 'Arial',
    fontSize: 14,
    labelSize: 14,
    textColor: '#111111',
    bgColor: '#ffffff',
    borderColor: '#e6e6ef',
    radius: 6,
    padding: 8,
    options: ['Option 1', 'Option 2'],
    colSpan: 1, // per-element span
    binding: { address: '', listen: false }
  }, data);

  wrapper._meta = meta;

  const labelEl = document.createElement('div');
  labelEl.className = 'field-label';
  labelEl.contentEditable = false;
  labelEl.textContent = meta.label;

  let fieldEl;
  switch(type) {
    case 'text':
    case 'email':
    case 'number':
    case 'password':
      fieldEl = document.createElement('input'); fieldEl.type = (type==='text'?'text':type); fieldEl.placeholder = meta.placeholder; break;
    case 'textarea':
      fieldEl = document.createElement('textarea'); fieldEl.placeholder = meta.placeholder; break;
    case 'select':
    case 'dropdown':
      fieldEl = document.createElement('select'); (meta.options||[]).forEach(o=>{const opt=document.createElement('option');opt.textContent=o;fieldEl.appendChild(opt);}); break;
    case 'checkbox':
      fieldEl = document.createElement('input'); fieldEl.type='checkbox'; break;
    case 'radio':
      fieldEl = document.createElement('div'); { const group='r'+Date.now(); (meta.options||[]).forEach(o=>{const r=document.createElement('label'); const inp=document.createElement('input'); inp.type='radio'; inp.name=group; r.appendChild(inp); r.appendChild(document.createTextNode(' '+o)); fieldEl.appendChild(r);}); } break;
    case 'button':
      fieldEl = document.createElement('button'); fieldEl.textContent = meta.label; break;
    case 'heading':
      fieldEl = document.createElement('div'); fieldEl.innerHTML = `<strong style="font-size:1.2em">${meta.label}</strong>`; break;
    case 'image':
      fieldEl = document.createElement('img'); fieldEl.src = data.src || ''; fieldEl.style.maxWidth='100%'; break;
    case 'date':
      fieldEl = document.createElement('input'); fieldEl.type = 'date'; break;
    case 'file':
      fieldEl = document.createElement('input'); fieldEl.type = 'file'; break;
    case 'color':
      fieldEl = document.createElement('input'); fieldEl.type = 'color'; break;
    case 'range':
      fieldEl = document.createElement('input'); fieldEl.type = 'range'; break;
    default:
      fieldEl = document.createElement('input'); fieldEl.type='text';
  }

  wrapper._control = fieldEl;
  wrapper.appendChild(labelEl);
  wrapper.appendChild(fieldEl);
  applyMetaToElement(wrapper);
  attachBindingHandlers(wrapper);

  on(wrapper, 'click', (e) => { e.stopPropagation(); selectElement(wrapper); });

  on(labelEl, 'dblclick', (e)=>{
    e.stopPropagation(); labelEl.contentEditable = true; labelEl.focus();
    const sel = window.getSelection(); sel.selectAllChildren(labelEl);
  });
  on(labelEl, 'blur', ()=>{
    labelEl.contentEditable = false; meta.label = labelEl.textContent; applyMetaToElement(wrapper); syncSettingsPanel(); autoSaveLayout();
  });

  // free-mode dragging
  const pos = {x:0,y:0,dragging:false};
  on(wrapper, 'pointerdown', (ev)=>{
    if(!isFreeMode) return; ev.preventDefault();
    wrapper.setPointerCapture(ev.pointerId);
    pos.dragging = true; pos.startX = ev.clientX; pos.startY = ev.clientY;
    pos.origLeft = parseInt(wrapper.style.left||0,10)||0; pos.origTop = parseInt(wrapper.style.top||0,10)||0;
  });
  on(wrapper, 'pointermove', (ev)=>{
    if(!isFreeMode || !pos.dragging) return;
    let dx=ev.clientX-pos.startX, dy=ev.clientY-pos.startY;
    let nx=pos.origLeft+dx, ny=pos.origTop+dy;
    if(snapToggle?.checked){ nx=Math.round(nx/gridSize)*gridSize; ny=Math.round(ny/gridSize)*gridSize; }
    wrapper.style.left = nx+'px'; wrapper.style.top = ny+'px';
  });
  on(wrapper, 'pointerup', (ev)=>{
    if(!isFreeMode) return; pos.dragging=false; try{ wrapper.releasePointerCapture(ev.pointerId);}catch(_){}
    autoSaveLayout();
  });

  // flow drag reorder
  wrapper.draggable = true;
  on(wrapper, 'dragstart', (ev)=>{
    if(isFreeMode){ ev.preventDefault(); return; }
    ev.dataTransfer.setData('text/field-id',''); wrapper.classList.add('dragging');
  });
  on(wrapper, 'dragend', ()=> { wrapper.classList.remove('dragging'); autoSaveLayout(); });

  return wrapper;
}

/* ---------- Options (select/radio) rebuild ---------- */
function rebuildOptions(wrapper){
  const type = wrapper.getAttribute('data-type');
  const meta = wrapper._meta;
  const ctl  = wrapper._control;
  if(!ctl) return;

  if(type === 'select' || type === 'dropdown'){
    ctl.innerHTML = '';
    (meta.options || []).forEach(o=>{
      const opt = document.createElement('option'); opt.textContent = o; ctl.appendChild(opt);
    });
  } else if(type === 'radio'){
    ctl.innerHTML = '';
    const groupName = 'r' + Date.now();
    (meta.options || []).forEach(o=>{
      const lbl = document.createElement('label');
      const inp = document.createElement('input'); inp.type='radio'; inp.name = groupName;
      lbl.appendChild(inp); lbl.appendChild(document.createTextNode(' '+o));
      ctl.appendChild(lbl);
    });
  }
}

/* ---------- Value helpers (binding uses these) ---------- */
function getControlValue(wrapper){
  const type = wrapper.getAttribute('data-type');
  const ctl  = wrapper._control;
  if(!ctl) return '';

  if(ctl.tagName === 'INPUT'){
    const t = ctl.type;
    if(['text','email','number','password','date','color','range','file'].includes(t)){
      if(t === 'file') { return ctl.files && ctl.files.length ? ctl.files[0].name : ''; }
      return ctl.value;
    }
    if(t === 'checkbox') return ctl.checked ? 'TRUE' : 'FALSE';
  }
  if(ctl.tagName === 'TEXTAREA') return ctl.value;
  if(ctl.tagName === 'SELECT') return ctl.value;

  if(type === 'radio'){
    const checked = ctl.querySelector('input[type="radio"]:checked');
    if(!checked) return '';
    const lbl = checked.parentElement;
    return (lbl && lbl.textContent) ? lbl.textContent.trim() : '';
  }

  if(type === 'button' || type === 'heading'){
    return wrapper._meta.label;
  }
  return '';
}

function setControlValue(wrapper, value){
  const type = wrapper.getAttribute('data-type');
  const ctl  = wrapper._control;
  if(!ctl) return;

  if(ctl.tagName === 'INPUT'){
    const t = ctl.type;
    if(['text','email','number','password','date','color','range'].includes(t)){ ctl.value = value ?? ''; return; }
    if(t === 'checkbox'){
      const v = (value || '').toString().toLowerCase();
      ctl.checked = (v === 'true' || v === '1' || v === 'yes' || v === 'on'); return;
    }
  }
  if(ctl.tagName === 'TEXTAREA'){ ctl.value = value ?? ''; return; }
  if(ctl.tagName === 'SELECT'){
    const match = Array.from(ctl.options).some(o=>o.value === value || o.text === value);
    if(!match && value != null){ const opt=document.createElement('option'); opt.textContent=value; ctl.appendChild(opt); }
    ctl.value = value ?? ''; return;
  }
  if(type === 'radio'){
    const want = (value ?? '').toString().trim();
    const labels = ctl.querySelectorAll('label');
    labels.forEach(l=>{
      const text = (l.textContent||'').trim();
      const input = l.querySelector('input[type="radio"]');
      if(input) input.checked = (text === want);
    });
  }
}

/* ---------- Apply meta to DOM ---------- */
function applyMetaToElement(wrapper){
  const m     = wrapper._meta;
  const label = wrapper.querySelector('.field-label');
  const input = wrapper._control;
  const type  = wrapper.getAttribute('data-type');

  label.textContent = m.label || '';
  if(input){
    if(input.tagName === 'INPUT' || input.tagName==='TEXTAREA'){
      input.placeholder = m.placeholder || '';
      input.required = !!m.required;
    }
    if(input.tagName === 'BUTTON') input.textContent = m.label || '';
    if(type === 'heading') input.innerHTML = `<strong style="font-size:1.2em">${m.label}</strong>`;

    if(type === 'select' || type === 'dropdown' || type === 'radio') rebuildOptions(wrapper);

    wrapper.style.width        = (m.widthPct ? m.widthPct + '%' : 'auto');
    if(m.height && m.height > 0) input.style.height = m.height + 'px'; else input.style.height = '';
    wrapper.style.padding      = m.padding + 'px';
    label.style.fontSize       = m.labelSize + 'px';
    label.style.fontFamily     = m.font;
    input.style.fontSize       = m.fontSize + 'px';
    input.style.fontFamily     = m.font;
    input.style.color          = m.textColor;
    input.style.background     = m.bgColor;
    input.style.borderColor    = m.borderColor;
    input.style.borderRadius   = m.radius + 'px';
  }

  // per-element grid span (flow mode only)
  if(!isFreeMode){
    const cols = getCurrentColumns();
    const span = clamp(Number(m.colSpan||1), 1, cols);
    m.colSpan = span;
    wrapper.style.gridColumn = `span ${span}`;
    wrapper.style.position   = 'relative';
    wrapper.style.left = wrapper.style.top = '';
    wrapper.style.width = '100%';
  }
}

/* ---------- Binding handlers (Shim or real BM.write/read) ---------- */
function detachBindingHandlers(wrapper){
  if(wrapper._boundInputTarget && wrapper._boundInputEvent && wrapper._boundInputHandler){
    wrapper._boundInputTarget.removeEventListener(wrapper._boundInputEvent, wrapper._boundInputHandler);
  }
  wrapper._boundInputTarget  = null;
  wrapper._boundInputEvent   = null;
  wrapper._boundInputHandler = null;

  if(wrapper._bindingSubscriptionId){
    try { BM.unsubscribe(wrapper._bindingSubscriptionId); } catch(_){}
  }
  wrapper._bindingSubscriptionId = null;
}

function attachBindingHandlers(wrapper){
  detachBindingHandlers(wrapper);
  const m   = wrapper._meta;
  const ctl = wrapper._control; if(!ctl) return;

  const type = wrapper.getAttribute('data-type');
  let target = ctl, eventName = 'input';
  if(ctl.tagName === 'SELECT') eventName = 'change';
  if(ctl.tagName === 'INPUT' && ['checkbox','color','date','range','file'].includes(ctl.type)) eventName = 'change';
  if(type === 'radio') { target = ctl; eventName = 'change'; }

  const handler = async () => {
    if(!m.binding || !m.binding.address) return;
    const val = getControlValue(wrapper);
    try { await BM.write(m.binding.address, val); } catch(e){ /* ignore */ }
  };
  target.addEventListener(eventName, handler);
  wrapper._boundInputTarget  = target;
  wrapper._boundInputEvent   = eventName;
  wrapper._boundInputHandler = handler;

  if(m.binding && m.binding.listen && m.binding.address && BM.isReady()){
    BM.ensureBinding(m.binding.address).then(bindingId=>{
      wrapper._bindingSubscriptionId = bindingId;
      BM.subscribe(bindingId, async ()=>{
        try{
          const v = await BM.read(m.binding.address);
          setControlValue(wrapper, v);
        }catch(_){}
      });
    }).catch(()=>{});
  }
}

/* ---------- Selection & Panels ---------- */
function clearSelection(){
  $$('.form-field').forEach(el => el.classList.remove('selected'));
  selectedEl = null;
  if(elementSettings) elementSettings.style.display='none';
  if(globalSettings)  globalSettings.style.display='block';
  if(selectedInfo)    selectedInfo.textContent = 'No element selected';
  if (tabGlobal) showPanel('global');
}

function selectElement(el){
  $$('.form-field').forEach(x=>x.classList.remove('selected'));
  el.classList.add('selected');
  selectedEl = el;
  if(elementSettings) elementSettings.style.display='block';
  if(globalSettings)  globalSettings.style.display='none';
  if(selectedInfo)    selectedInfo.textContent = `${el._meta.label} — ${el.getAttribute('data-type')}`;
  if (tabElement) showPanel('element');
  syncSettingsPanel();
}

function syncSettingsPanel(){
  if(!selectedEl) return;
  const m = selectedEl._meta;

  if(elLabel)       elLabel.value       = m.label || '';
  if(elPlaceholder) elPlaceholder.value = m.placeholder || '';
  if(elRequired)    elRequired.checked  = !!m.required;
  if(elWidth)       elWidth.value       = m.widthPct;
  if(elHeight)      elHeight.value      = m.height || 0;
  if(elFont)        elFont.value        = m.font;
  if(elFontSize)    elFontSize.value    = m.fontSize;
  if(elLabelSize)   elLabelSize.value   = m.labelSize;
  if(elTextColor)   elTextColor.value   = m.textColor;
  if(elBgColor)     elBgColor.value     = m.bgColor;
  if(elBorderColor) elBorderColor.value = m.borderColor;
  if(elRadius)      elRadius.value      = m.radius;
  if(elPadding)     elPadding.value     = m.padding;

  if(elColSpan){
    const cols = getCurrentColumns();
    elColSpan.min = 1;
    elColSpan.max = Math.max(1, cols);
    elColSpan.value = clamp(Number(m.colSpan||1), 1, cols);
  }

  if(elBindAddress) elBindAddress.value = (m.binding && m.binding.address) ? m.binding.address : '';
  if(elBindListen)  elBindListen.checked= !!(m.binding && m.binding.listen);

  const type = selectedEl.getAttribute('data-type');
  if(optionsEditor){
    if(['select','dropdown','radio'].includes(type)){ optionsEditor.style.display='block'; renderOptionsList(); }
    else { optionsEditor.style.display='none'; }
  }
}

function applySettingsFromPanel(){
  if(!selectedEl) return;
  const m = selectedEl._meta;

  if(elLabel)       m.label       = elLabel.value;
  if(elPlaceholder) m.placeholder = elPlaceholder.value;
  if(elRequired)    m.required    = elRequired.checked;
  if(elWidth)       m.widthPct    = Number(elWidth.value)||100;
  if(elHeight)      m.height      = Number(elHeight.value)||0;
  if(elFont)        m.font        = elFont.value;
  if(elFontSize)    m.fontSize    = Number(elFontSize.value)||14;
  if(elLabelSize)   m.labelSize   = Number(elLabelSize.value)||14;
  if(elTextColor)   m.textColor   = elTextColor.value;
  if(elBgColor)     m.bgColor     = elBgColor.value;
  if(elBorderColor) m.borderColor = elBorderColor.value;
  if(elRadius)      m.radius      = Number(elRadius.value)||0;
  if(elPadding)     m.padding     = Number(elPadding.value)||8;

  if(elColSpan){
    const cols = getCurrentColumns();
    m.colSpan = clamp(Number(elColSpan.value)||1, 1, cols);
  }

  applyMetaToElement(selectedEl);
  autoSaveLayout();
}

/* ---------- Options Editor ---------- */
function renderOptionsList(){
  if(!selectedEl || !optionsList) return;
  optionsList.innerHTML='';
  const opts = selectedEl._meta.options || [];
  opts.forEach((o,i)=>{
    const row = document.createElement('div'); row.className='opt-row';
    const txt = document.createElement('input'); txt.value = o;
    on(txt, 'input', ()=>{
      selectedEl._meta.options[i]=txt.value;
      applyMetaToElement(selectedEl);
      autoSaveLayout();
    });
    const del = document.createElement('button'); del.textContent='✕'; del.className='btn ghost';
    on(del, 'click', ()=>{
      selectedEl._meta.options.splice(i,1);
      renderOptionsList();
      applyMetaToElement(selectedEl);
      autoSaveLayout();
    });
    row.appendChild(txt); row.appendChild(del); optionsList.appendChild(row);
  });
}
on(addOptionBtn, 'click', ()=>{
  if(!selectedEl || !newOptionText) return;
  const v = newOptionText.value.trim(); if(!v) return;
  selectedEl._meta.options = selectedEl._meta.options || []; selectedEl._meta.options.push(v);
  newOptionText.value='';
  renderOptionsList();
  applyMetaToElement(selectedEl);
  autoSaveLayout();
});

/* ---------- Toolbox drag/drop ---------- */
toolbox.forEach(t=>{
  t.setAttribute('draggable','true');
  on(t, 'dragstart', e=>{
    dragDataType = t.dataset.type;
    e.dataTransfer.setData('text/plain','tool');
  });
});
on(canvas, 'dragover', e=>{ e.preventDefault(); });
on(canvas, 'drop', e=>{
  e.preventDefault();
  if(!dragDataType) return;
  const newField = createField(dragDataType);
  if(layoutMode.value === 'flow'){
    formRoot.appendChild(newField);
  } else {
    const rect = formRoot.getBoundingClientRect();
    newField.style.position = 'absolute';
    let left = e.clientX - rect.left;
    let top  = e.clientY - rect.top;
    if(snapToggle?.checked){
      left = Math.round(left / gridSize) * gridSize;
      top  = Math.round(top  / gridSize) * gridSize;
    }
    newField.style.left = left + 'px';
    newField.style.top  = top  + 'px';
    formRoot.appendChild(newField);
  }
  const empty = formRoot.querySelector('.empty-state'); if(empty) empty.remove();
  selectElement(newField);
  dragDataType = null;
  autoSaveLayout();
});

/* ---------- Flow reorder ---------- */
on(formRoot, 'dragover', e=>{
  if(layoutMode.value !== 'flow') return;
  e.preventDefault();
  const dragging = document.querySelector('.form-field.dragging');
  const afterEl = getDragAfterElement(formRoot, e.clientY);
  if(dragging){
    if(afterEl == null) formRoot.appendChild(dragging);
    else formRoot.insertBefore(dragging, afterEl);
  }
});
function getDragAfterElement(container, y){
  const els = [...container.querySelectorAll('.form-field:not(.dragging)')];
  return els.reduce((closest, child)=>{
    const box = child.getBoundingClientRect();
    const offset = y - box.top - box.height/2;
    if(offset < 0 && offset > closest.offset) return {offset, element: child};
    return closest;
  }, {offset: Number.NEGATIVE_INFINITY}).element;
}

/* ---------- Click blank to show Global ---------- */
on(canvas, 'click', (e)=>{
  if(e.target === canvas || e.target === formRoot || e.target.classList.contains('empty-state')) {
    clearSelection();
  }
});

/* ---------- Bind settings inputs ---------- */
[
  elLabel, elPlaceholder, elRequired, elWidth, elHeight, elFont, elFontSize,
  elLabelSize, elTextColor, elBgColor, elBorderColor, elRadius, elPadding, elColSpan
].forEach(inp => {
  if (!inp) return;
  const evt = (inp.type === 'checkbox' || inp.type === 'radio') ? 'change' : 'input';
  on(inp, evt, ()=> applySettingsFromPanel());
});

/* ---------- Binding UI events (enhanced with BindingManager) ---------- */
on(elBindAddress, 'change', ()=>{
  if(!selectedEl) return;
  const addr = elBindAddress.value.trim();
  selectedEl._meta.binding = selectedEl._meta.binding || {address:'', listen:false};
  selectedEl._meta.binding.address = addr;

  if (window.BindingManager && typeof BindingManager.bindControlToRange === 'function' && addr) {
    BindingManager.bindControlToRange(selectedEl._meta.id, addr);
  } else {
    attachBindingHandlers(selectedEl);
  }
  autoSaveLayout();
});

on(elBindListen, 'change', ()=>{
  if(!selectedEl) return;
  selectedEl._meta.binding = selectedEl._meta.binding || {address:'', listen:false};
  selectedEl._meta.binding.listen = !!elBindListen.checked;

  const addr = selectedEl._meta.binding.address;
  if (elBindListen.checked && addr && window.BindingManager && typeof BindingManager.bindControlToRange === 'function') {
    BindingManager.bindControlToRange(selectedEl._meta.id, addr);
  } else {
    attachBindingHandlers(selectedEl);
  }
  autoSaveLayout();
});

on(bindWriteBtn, 'click', async ()=>{
  if(!selectedEl) return;
  const m = selectedEl._meta;
  if(!m.binding || !m.binding.address) return alert('Set a binding address first.');
  const val = getControlValue(selectedEl);
  try { await BM.write(m.binding.address, val); } catch(_){}
});

on(bindReadBtn, 'click', async ()=>{
  if(!selectedEl) return;
  const m = selectedEl._meta;
  if(!m.binding || !m.binding.address) return alert('Set a binding address first.');
  try { const v = await BM.read(m.binding.address); setControlValue(selectedEl, v); } catch(_){}
});

/* ---------- Global settings ---------- */
on(formBgColor, 'input', ()=>{ formRoot.style.background = formBgColor.value || ''; autoSaveLayout(); });

let _bgImageDataUrl = '';
on(formBgImage, 'change', ()=>{
  const f = formBgImage?.files && formBgImage.files[0];
  if(!f){ _bgImageDataUrl=''; formRoot.style.backgroundImage=''; autoSaveLayout(); return; }
  const reader = new FileReader();
  reader.onload = e=>{
    _bgImageDataUrl = e.target.result || '';
    formRoot.style.backgroundImage = _bgImageDataUrl ? `url("${_bgImageDataUrl}")` : '';
    formRoot.style.backgroundSize = 'cover';
    formRoot.style.backgroundRepeat = 'no-repeat';
    formRoot.style.backgroundPosition = 'center';
    autoSaveLayout();
  };
  reader.readAsDataURL(f);
});

on(formWidth, 'input', ()=>{ formRoot.style.maxWidth = (Number(formWidth.value)||0) ? (formWidth.value + 'px') : ''; autoSaveLayout(); });
on(formPadding, 'input', ()=>{ formRoot.style.padding = (Number(formPadding.value)||0) ? (formPadding.value + 'px') : ''; autoSaveLayout(); });

on(gridToggle, 'change', ()=> { toggleGrid(gridToggle.checked); autoSaveLayout(); });
on(gridSizeInput, 'input', ()=>{ gridSize = parseInt(gridSizeInput.value,10)||12; updateGrid(); autoSaveLayout(); });

/* ---------- Columns (flow mode) ---------- */
on(columnsSelect, 'change', ()=>{
  applyLayoutColumns();
  $$('.form-field').forEach(el=>{
    const m = el._meta || {};
    const cols = getCurrentColumns();
    m.colSpan = clamp(Number(m.colSpan||1), 1, cols);
    if (el === selectedEl && elColSpan) elColSpan.value = m.colSpan;
    applyMetaToElement(el);
  });
  autoSaveLayout();
});

function applyLayoutColumns(){
  if(layoutMode.value !== 'flow') return;
  const cols = getCurrentColumns();
  formRoot.style.display = 'grid';
  formRoot.style.gridTemplateColumns = `repeat(${cols}, 1fr)`;
  if(!formRoot.style.gap) formRoot.style.gap = '12px';
  $$('.form-field').forEach(el=>{
    const m = el._meta || {};
    const span = clamp(Number(m.colSpan||1), 1, cols);
    m.colSpan = span;
    el.style.gridColumn = `span ${span}`;
    el.style.position = 'relative';
    el.style.left = el.style.top = '';
    el.style.width = '100%';
  });
}

/* ---------- Grid overlay ---------- */
function toggleGrid(show){
  if(!gridOverlay) return;
  gridOverlay.style.opacity = show ? 1 : 0;
  gridOverlay.style.pointerEvents = 'none';
  gridOverlay.style.backgroundSize = `${gridSize}px ${gridSize}px, ${gridSize}px ${gridSize}px`;
}
function updateGrid(){
  if(!gridOverlay) return;
  gridOverlay.style.backgroundSize = `${gridSize}px ${gridSize}px, ${gridSize}px ${gridSize}px`;
}
updateGrid();

/* ---------- Layout mode ---------- */
on(layoutMode, 'change', ()=>{
  isFreeMode = (layoutMode.value === 'free');
  if(!isFreeMode){
    formRoot.style.display = 'grid';
    $$('.form-field').forEach(el=>{
      el.style.position='relative'; el.style.left=''; el.style.top=''; el.style.zIndex='';
      applyMetaToElement(el);
    });
    applyLayoutColumns();
  } else {
    formRoot.style.display = 'block';
    const rect = formRoot.getBoundingClientRect();
    $$('.form-field').forEach(el=>{
      const r = el.getBoundingClientRect();
      el.style.position='absolute';
      el.style.left = (r.left - rect.left) + 'px';
      el.style.top  = (r.top  - rect.top)  + 'px';
      el.style.width = (el._meta?.widthPct ? el._meta.widthPct + '%' : 'auto');
      el.style.gridColumn = '';
    });
  }
  autoSaveLayout();
});

/* ---------- Export / Import ---------- */
function exportLayout(){
  const nodes = [];
  $$('.form-field').forEach(n=>{
    const meta = Object.assign({}, n._meta);
    // Ensure both colSpan and span for cross-version compatibility
    if (typeof meta.colSpan !== 'undefined') meta.span = meta.colSpan;

    nodes.push({
      type: n.getAttribute('data-type'),
      meta: meta,
      pos: { left: n.style.left || null, top: n.style.top || null }
    });
  });
  const payload = {
    global: {
      bg: formBgColor?.value || '',
      bgImageDataUrl: _bgImageDataUrl || '',
      width: formWidth?.value || '',
      padding: formPadding?.value || '',
      layout: layoutMode?.value || 'flow',
      columns: columnsSelect?.value || '1',
      grid: !!gridToggle?.checked,
      gridSize: gridSize
    },
    nodes
  };
  return JSON.stringify(payload, null, 2);
}

on(exportJsonBtn, 'click', ()=>{
  const json = exportLayout();
  const a = document.createElement('a');
  a.href = 'data:application/json;charset=utf-8,'+encodeURIComponent(json);
  a.download = 'form_layout.json';
  a.click();
});

on(importJsonBtn, 'click', ()=> importFileInput && importFileInput.click());
on(importFileInput, 'change', (ev)=>{
  const f = ev.target.files[0]; if(!f) return;
  const reader = new FileReader();
  reader.onload = e => { loadFromJson(e.target.result); autoSaveLayout(); }
  reader.readAsText(f);
});

/* Save button -> Excel hidden sheet */
on(saveLayoutBtn, 'click', ()=>{
  saveLayoutToExcel();
});

/* Hidden sheet persistence (Excel) */
async function saveLayoutToExcel() {
  if (!(window.Excel && Excel.run)) { console.warn('Excel API not available; skipping save.'); return; }
  try {
    showSaveStatus("Saving...");
    await Excel.run(async (context) => {
      const wb = context.workbook;
      let sheet;
      try { sheet = wb.worksheets.getItem("_FormDesignerMeta"); }
      catch { sheet = wb.worksheets.add("_FormDesignerMeta"); }

      sheet.getRange("A1").clear();
      const json = exportLayout();
      sheet.getRange("A1").values = [[json]];
      sheet.visibility = Excel.SheetVisibility.hidden;

      await context.sync();
      showSaveStatus("All changes saved");
      console.log("✅ Layout saved to hidden sheet");
    });
  } catch (err) {
    showSaveStatus("Save failed!");
    console.error("❌ Failed to save layout to Excel:", err);
  }
}

async function loadLayoutFromExcel() {
  if (!(window.Excel && Excel.run)) { console.warn('Excel API not available; skipping load.'); return; }
  try {
    await Excel.run(async (context) => {
      const wb = context.workbook;
      let sheet;
      try { sheet = wb.worksheets.getItem("_FormDesignerMeta"); }
      catch (e) { console.log("ℹ️ No saved layout sheet found."); return; }

      const rng = sheet.getRange("A1").load("values");
      await context.sync();

      const raw = rng.values[0][0];
      if (raw && raw.trim()) {
        try { loadFromJson(raw); console.log("✅ Layout loaded from hidden sheet"); }
        catch (err) { console.error("❌ Failed to parse saved layout JSON:", err); }
      }
    });
  } catch (err) {
    console.error("❌ Failed to load layout from Excel:", err);
  }
}

/* Debounced autosave trigger (2s) */
const autoSaveLayout = debounce(() => { saveLayoutToExcel(); }, 2000);

/* ---------- Load from JSON (with span→colSpan Option B) ---------- */
function loadFromJson(str){
  try{
    const obj = JSON.parse(str);

    // global -> inputs
    if(formBgColor) formBgColor.value = obj.global.bg || '';
    _bgImageDataUrl = obj.global.bgImageDataUrl || '';
    if(formBgImage) formBgImage.value = ''; // cannot prefill file inputs
    if(_bgImageDataUrl){
      formRoot.style.backgroundImage = `url("${_bgImageDataUrl}")`;
      formRoot.style.backgroundSize = 'cover';
      formRoot.style.backgroundRepeat = 'no-repeat';
      formRoot.style.backgroundPosition = 'center';
    } else {
      formRoot.style.backgroundImage = '';
    }

    if(formWidth)   formWidth.value   = obj.global.width   || 900;
    if(formPadding) formPadding.value = obj.global.padding || 16;
    if(layoutMode)  layoutMode.value  = obj.global.layout  || 'flow';
    if(columnsSelect) columnsSelect.value = obj.global.columns || '1';
    gridSize = parseInt(obj.global.gridSize,10) || gridSize;
    if(gridSizeInput) gridSizeInput.value = gridSize;
    if(gridToggle) gridToggle.checked = !!obj.global.grid;

    // apply global styles
    formRoot.style.background = formBgColor.value || '';
    formRoot.style.maxWidth   = formWidth.value ? (formWidth.value + 'px') : '';
    formRoot.style.padding    = formPadding.value ? (formPadding.value + 'px') : '';
    updateGrid(); toggleGrid(gridToggle ? gridToggle.checked : false);

    // rebuild nodes
    formRoot.innerHTML = '';
    if (obj.nodes && Array.isArray(obj.nodes)) {
      obj.nodes.forEach(n=>{
        // Option B: accept V2 'span' by mapping to V1 'colSpan' if needed
        if (n.meta && n.meta.span && !n.meta.colSpan) {
          n.meta.colSpan = n.meta.span;
        }
        const f = createField(n.type, n.meta || {});
        if(n.pos && n.pos.left && (obj.global.layout === 'free')){
          f.style.position='absolute';
          f.style.left = n.pos.left;
          f.style.top  = n.pos.top;
        }
        formRoot.appendChild(f);
      });
    }

    // set layout
    isFreeMode = (layoutMode.value === 'free');
    if (isFreeMode) {
      formRoot.style.display = 'block';
    } else {
      applyLayoutColumns();
    }
  }catch(err){
    alert('Invalid JSON');
  }
}

/* ---------- Clear, Bring front, Delete ---------- */
on(clearCanvasBtn, 'click', ()=>{
  if(!confirm('Clear the canvas and remove all fields?')) return;
  formRoot.innerHTML = '<div class="empty-state">Drop elements here — this is your form canvas.</div>';
  clearSelection();
  autoSaveLayout();
});

on($('#bringFront'), 'click', ()=>{
  if(!selectedEl) return;
  selectedEl.style.zIndex = (Number(selectedEl.style.zIndex)||0) + 1;
  autoSaveLayout();
});

on($('#deleteElement'), 'click', ()=>{
  if(!selectedEl) return;
  if(!confirm('Delete selected element?')) return;
  detachBindingHandlers(selectedEl);
  selectedEl.remove();
  clearSelection();
  autoSaveLayout();
});

/* ---------- Inline label live sync ---------- */
on(document, 'input', (e)=>{
  if(!selectedEl) return;
  const lbl = e.target.closest('.field-label');
  if(lbl && selectedEl.contains(lbl)){
    selectedEl._meta.label = lbl.textContent;
    if(selectedInfo) selectedInfo.textContent = selectedEl._meta.label;
    applyMetaToElement(selectedEl);
    autoSaveLayout();
  }
});

/* ---------- ESC to clear ---------- */
on(document, 'keydown', (e)=>{ if(e.key==='Escape') clearSelection(); });

/* ---------- Optional tabs ---------- */
function showPanel(which){
  if(!tabGlobal || !tabElement || !elementSettings || !globalSettings) return;
  if(which === 'element'){
    tabElement.classList.add('active'); tabGlobal.classList.remove('active');
    elementSettings.style.display='block'; globalSettings.style.display='none';
  } else {
    tabGlobal.classList.add('active'); tabElement.classList.remove('active');
    elementSettings.style.display='none'; globalSettings.style.display='block';
  }
}
on(tabGlobal,  'click', ()=> showPanel('global'));
on(tabElement, 'click', ()=> { if(selectedEl) showPanel('element'); });

/* ---------- Initial grid + columns ---------- */
toggleGrid(gridToggle ? gridToggle.checked : false);
applyLayoutColumns();

/* ---------- Office onReady: load from hidden sheet ---------- */
if (window.Office && Office.onReady) {
  Office.onReady(() => {
    loadLayoutFromExcel();
  });
}

/* ---------- Demo seed (can be removed) ---------- */
(function initDemo(){
  const f1 = createField('heading', {label:'Customer Signup', colSpan: getCurrentColumns()});
  const f2 = createField('text',   {label:'Full name', placeholder:'Jane Doe', colSpan: 1});
  const f3 = createField('email',  {label:'Email address', placeholder:'name@example.com', colSpan: 1});
  formRoot.appendChild(f1); formRoot.appendChild(f2); formRoot.appendChild(f3);
})();
