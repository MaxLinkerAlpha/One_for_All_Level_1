const DATA_LOOM = {{DATA_LOOM}};
const DATA_LLPSI = {{DATA_LLPSI}};
const CATEGORIES = {{CATEGORIES}};
const SHEET3_ROOTS = {{SHEET3_ROOTS}};
const LATIN_PHRASES = {{LATIN_PHRASES}};
const LOANWORDS = {{LOANWORDS}};
const MYTHOLOGY = {{MYTHOLOGY}};
const BIBLE = {{BIBLE}};
const ENCODING = {{ENCODING}};
const FORVO_LANGS={latin:'la',english:'en',french:'fr',spanish:'es',portuguese:'pt',italian:'it',swedish:'sv',danish:'da',dutch:'nl',german:'de',russian:'ru',polish:'pl',grc:'grc',lat:'la'};
const SPECIAL_COL_MAP = {
  s3roots: [
    {key:'s3_greek',label:'希腊语'},
    {key:'s3_transliteration',label:'音译'},
    {key:'s3_english_meaning',label:'英文释义'},
    {key:'s3_latin_root',label:'拉丁词根'},
    {key:'s3_examples',label:'示例'},
    {key:'s3_section',label:'分组'}
  ],
  phrases: [
    {key:'phr_latin',label:'拉丁语'},
    {key:'phr_english',label:'英语'},
    {key:'phr_chinese',label:'中文'}
  ],
  loanwords: [
    {key:'lw_word',label:'单词'},
    {key:'lw_english_note',label:'英文说明'},
    {key:'lw_chinese_note',label:'中文说明'},
    {key:'lw_example',label:'示例'}
  ],
  mythology: [
    {key:'myth_greek',label:'希腊名'},
    {key:'myth_latin',label:'拉丁名'},
    {key:'myth_greek_deity',label:'希腊神'},
    {key:'myth_roman_deity',label:'罗马神'},
    {key:'myth_planet',label:'行星'},
    {key:'myth_notes',label:'备注'}
  ],
  bible: [
    {key:'bib_book',label:'经文'},
    {key:'bib_latin',label:'拉丁原文'},
    {key:'bib_chinese',label:'中文翻译'},
    {key:'bib_keyword',label:'关键词'},
    {key:'bib_peg',label:'桩内容'},
    {key:'bib_mnemonic',label:'助记'}
  ],
  encoding: [
    {key:'enc_number',label:'数字'},
    {key:'enc_cardinal',label:'基数词'},
    {key:'enc_ordinal',label:'序数词'},
    {key:'enc_name',label:'人名'},
    {key:'enc_action',label:'动作'},
    {key:'enc_item',label:'物品'},
    {key:'enc_story',label:'故事'}
  ]
};

const COL_DEFS = [
  {key:'latin',label:'\u62C9\u4E01\u8BED',group:'lang',tabs:['all','romance','germanic','slavic','phrases']},
  {key:'greek_root',label:'\u5E0C\u814A\u8BCD\u6839',group:'lang',tabs:['all','romance','germanic','slavic']},
  {key:'english',label:'\u82F1\u8BED',group:'lang',tabs:['all','romance','germanic','slavic','phrases','loanwords']},
  {key:'french',label:'\u6CD5\u8BED',group:'lang',tabs:['all','romance','slavic']},
  {key:'spanish',label:'\u897F\u8BED',group:'lang',tabs:['all','romance']},
  {key:'portuguese',label:'\u8461\u8BED',group:'lang',tabs:['all','romance']},
  {key:'italian',label:'\u610F\u8BED',group:'lang',tabs:['all','romance']},
  {key:'swedish',label:'\u745E\u5178\u8BED',group:'lang',tabs:['all','germanic']},
  {key:'danish',label:'\u4E39\u8BED',group:'lang',tabs:['all','germanic']},
  {key:'dutch',label:'\u8377\u8BED',group:'lang',tabs:['all','germanic']},
  {key:'german',label:'\u5FB7\u8BED',group:'lang',tabs:['all','germanic','slavic']},
  {key:'russian',label:'\u4FC4\u8BED',group:'lang',tabs:['all','slavic']},
  {key:'polish',label:'\u6CE2\u5170\u8BED',group:'lang',tabs:['all','slavic']},
  {key:'category',label:'\u5206\u7C7B',group:'info',tabs:['all','romance','germanic','slavic']},
  {key:'pos',label:'\u8BCD\u6027',group:'info',tabs:['all','romance','germanic','slavic']},
  {key:'chinese_translation',label:'\u6C49\u8BED',group:'info',tabs:['all','romance','germanic','slavic','phrases','loanwords']},
  {key:'notes',label:'\u6CE8\u91CA',group:'info',tabs:['all','romance','germanic','slavic']},
  {key:'example_sentence',label:'\u4F8B\u53E5',group:'info',tabs:['all','romance','germanic','slavic','loanwords']},
  {key:'mnemonic_liu',label:'\u5229\u739B\u7AA6\u52A9\u8BB0',group:'info',tabs:['all','romance','germanic','slavic']},
  {key:'etymology',label:'\u8BCD\u6839',group:'info',tabs:['all','romance','germanic','slavic']}
];

var state = {
  tab:'all',query:'',category:'',
  colVisible:{{COL_VISIBLE}},
  page:0,filteredCache:null,activeRootKey:null,
  s3Page:0,phrasePage:0,loanwordPage:0,
  mythPage:0,biblePage:0,encPage:0,
  dataSource:'loom',
  pageSize:100
};

var fullTextStore={};
var fullTextCounter=0;

function esc(s){if(s==null)return '';return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function escJs(s){if(s==null)return '';return String(s).replace(/\\/g,'\\\\').replace(/'/g,"\\'").replace(/"/g,'\\"').replace(/\n/g,'\\n');}

function getDATA(){return state.dataSource==='llpsi'?DATA_LLPSI:DATA_LOOM;}

function forvoUrl(word,langCode){
  if(!word)return '';
  var w=word.replace(/\s*[(\uff08].*$/,'').trim();
  if(!w)return '';
  var base='https://zh.forvo.com/search/'+encodeURIComponent(w)+'/';
  if(langCode)base+=langCode+'/';
  return base;
}
function forvoBtn(word,lang){
  if(!word)return '';
  var w=word.replace(/\s*[(\uff08].*$/,'').trim();
  if(!w||w.length<1)return '';
  var lc=FORVO_LANGS[lang]||'';
  return '<a class="forvo-btn" href="'+forvoUrl(w,lc)+'" target="_blank" rel="noopener" title="\u53D1\u97F3" onclick="event.stopPropagation()">&#9835;</a>';
}

function init(){
  var saved=localStorage.getItem('theme');
  var isDark=saved==='dark'||(!saved&&window.matchMedia('(prefers-color-scheme:dark)').matches);
  if(isDark)document.documentElement.setAttribute('data-theme','dark');
  updateThemeIcon(isDark);
  buildRootMap();renderHeader();renderCards(true);
  document.getElementById('themeBtn').addEventListener('click',toggleTheme);
  document.getElementById('cardList').addEventListener('click',function(e){
    var el=e.target.closest('.grid-cell-text[data-ft]');
    if(el)showPopover(el);
  });
  document.addEventListener('keydown',function(e){
    var searchInput=document.getElementById('vocabSearch');
    if(e.key==='/'&&document.activeElement!==searchInput&&!e.ctrlKey&&!e.metaKey){e.preventDefault();searchInput.focus();}
    if(e.key==='Escape'){
      if(state.activeRootKey){clearRootFilter();return;}
      if(searchInput&&searchInput.value){searchInput.value='';state.query='';var cl=document.getElementById('searchClear');if(cl)cl.classList.remove('visible');resetAndRender();searchInput.blur();}
    }
  });
}

function resetAndRender(){state.page=0;state.filteredCache=null;state.s3Page=0;state.phrasePage=0;state.loanwordPage=0;state.mythPage=0;state.biblePage=0;state.encPage=0;renderCards(true);}
function toggleTheme(){var h=document.documentElement;var d=h.getAttribute('data-theme')==='dark';h.setAttribute('data-theme',d?'light':'dark');localStorage.setItem('theme',d?'light':'dark');updateThemeIcon(!d);}
function updateThemeIcon(d){var b=document.getElementById('themeBtn');if(b)b.textContent=d?'\u2600':'\u263E';}

function buildRootMap(){
  window.rootMap={};
  DATA_LOOM.forEach(function(item){
    var ll=item.latin||item.lingua_latina;
    if(ll&&ll.value){
      var r=ll.value.split(',')[0].trim().toLowerCase();
      if(!window.rootMap[r])window.rootMap[r]=[];
      window.rootMap[r].push(item.english+'|'+item.category);
    }
  });
  window.rootMapS3={};
  SHEET3_ROOTS.forEach(function(r){
    if(r.latin_root){
      var k=r.latin_root.toLowerCase().trim();
      if(!window.rootMapS3[k])window.rootMapS3[k]=[];
      window.rootMapS3[k].push(r);
    }
  });
}

function getLangCols(tab){
  if(tab==='romance')return['latin','greek_root','english','french','spanish','portuguese','italian'];
  if(tab==='germanic')return['latin','greek_root','english','swedish','danish','dutch','german'];
  if(tab==='slavic')return['latin','greek_root','english','french','german','russian','polish'];
  if(tab==='phrases')return['latin','english'];
  return['latin','greek_root','english','french','spanish','portuguese','italian','swedish','danish','dutch','german','russian','polish'];
}
function _colEmpty(col,data){
  if(!data||!data.length)return true;
  for(var i=0;i<Math.min(data.length,200);i++){
    var v=data[i][col];
    if(v){
      if(typeof v==='object'&&v.value&&v.value.trim())return false;
      if(typeof v==='string'&&v.trim())return false;
    }
  }
  return true;
}
function getVisibleLangs(tab){
  var cols=getLangCols(tab);
  var DATA=getDATA();
  var isLLPSI=state.dataSource==='llpsi';
  var src=DATA;
  if(isLLPSI){
    var items=DATA;
    cols=cols.filter(function(l){
      if(state.colVisible[l]===false)return false;
      return !_colEmpty(l,items);
    });
    return cols;
  }
  return cols.filter(function(l){return state.colVisible[l];});
}
function getVisibleInfo(){
  var cols=['category','pos','chinese_translation','notes','example_sentence','mnemonic_liu','etymology'];
  var DATA=getDATA();
  var isLLPSI=state.dataSource==='llpsi';
  if(isLLPSI){
    return cols.filter(function(k){return state.colVisible[k]!==false&&!_colEmpty(k,DATA);});
  }
  return cols.filter(function(k){return state.colVisible[k];});
}

function getInfoContent(item,info){
  if(info==='category')return item.category||'';
  if(info==='pos'){var p=item.pos;if(!p)return '';var labels={noun:'\u540D\u8BCD',verb:'\u52A8\u8BCD',adjective:'\u5F62\u5BB9\u8BCD',other:'\u5176\u4ED6'};return labels[p]||p;}
  if(info==='chinese_translation'){var f=item.chinese_translation;return f&&f.value?f.value:'';}
  if(info==='notes'){var f=item.notes;return f&&f.value?f.value:'';}
  if(info==='example_sentence'){var f=item.example_sentence;return f&&f.value?f.value:'';}
  if(info==='mnemonic_liu'){var f=item.mnemonic_liu;return f&&f.value?f.value:'';}
  if(info==='etymology'){return renderEtymologyText(item);}
  return '';
}
function renderEtymologyText(item){
  var h='';
  if(item.etymology&&item.etymology.length){
    item.etymology.forEach(function(e){
      var m=e.greek_meaning||e.english_meaning||'';
      var s=e.section||'';
      var x=e.english_examples_lat||e.english_examples_eng||'';
      var t=e.transliteration||'';
      if(!m&&!s&&!x)return;
      h+=s?'['+esc(s)+'] ':'';
      h+=m?esc(m):'';
      h+=t?(' ('+esc(t)+')'):'';
      h+=x?(' \u2014 '+esc(x)):'';
      h+='; ';
    });
  }
  if(item.llpsi_definition){
    var l=item.llpsi_definition;
    var ds=[l.english,l.french,l.spanish,l.italian,l.german].filter(Boolean);
    if(ds.length)h+='LLPSI: '+ds.map(function(d){return esc(d)}).join(', ');
  }
  return h.replace(/; $/,'');
}

function filterItems(){
  if(state.filteredCache)return state.filteredCache;
  var DATA=getDATA();
  var isLLPSI=state.dataSource==='llpsi';
  var r=DATA.filter(function(item){
    if(state.tab!=='all'&&['s3roots','phrases','loanwords','mythology','bible','encoding'].indexOf(state.tab)===-1){
      if(isLLPSI){
        if(state.tab==='romance')return true;
        if(state.tab==='germanic')return true;
        if(state.tab==='slavic')return true;
        return false;
      }
      var fam=item.family||'';
      // 'both' and 'both+slavic' contain both romance and germanic data
      var isBoth=fam==='both'||fam.indexOf('both')!==-1;
      var matchTab=fam===state.tab||(fam.indexOf(state.tab)!==-1);
      if(!matchTab&&!isBoth)return false;
    }
    if(state.category&&item.category!==state.category)return false;
    if(state.query){
      var vals=[item.english,item.french,item.spanish,item.portuguese,item.italian,item.swedish,item.danish,item.dutch,item.german,item.russian,item.polish];
      var lat=item.latin||item.lingua_latina;if(lat){vals.push(typeof lat==='object'?lat.value:lat);}
      var t=vals.filter(Boolean).join(' ');
      if(t.toLowerCase().indexOf(state.query)===-1)return false;
    }
    if(state.activeRootKey){
      var ll=item.latin||item.lingua_latina;
      if(!ll)return false;
      var lv=typeof ll==='object'?ll.value:ll;
      if(!lv)return false;
      if(lv.split(',')[0].trim().toLowerCase()!==state.activeRootKey)return false;
    }
    return true;
  });
  state.filteredCache=r;return r;
}
function gf(arr,fields){if(!state.query)return arr;var q=state.query.toLowerCase();return arr.filter(function(r){return fields.some(function(f){return r[f]&&String(r[f]).toLowerCase().indexOf(q)!==-1;});});}

function renderHeader(){
  var tabsEl=document.getElementById('tabs');
  var controlsEl=document.getElementById('headerControls');
  tabsEl.innerHTML='';
  controlsEl.innerHTML='';

  [['all','\u5168\u90E8'],['romance','\u7F57\u66FC'],['germanic','\u65E5\u8033\u66FC'],['slavic','\u65AF\u62C9\u592B'],['s3roots','\u8BCD\u6839\u8BCD\u7F00'],['phrases','\u62C9\u4E01\u77ED\u8BED'],['loanwords','\u501F\u8BCD'],['mythology','\u795E\u8BDD'],['bible','\u5723\u7ECF\u52A9\u8BB0'],['encoding','\u6570\u5B57\u7F16\u7801']].forEach(function(pair){
    var t=pair[0],l=pair[1];
    var b=document.createElement('button');b.className='tab'+(state.tab===t?' active':'');b.textContent=l;
    b.onclick=function(){state.tab=t;state.activeRootKey=null;updateRootBar();renderHeader();resetAndRender();};
    tabsEl.appendChild(b);
  });

  var special=['s3roots','phrases','loanwords','mythology','bible','encoding'];
  var isSpecial=special.indexOf(state.tab)!==-1;

  var searchWrap=document.createElement('div');searchWrap.className='search-wrap';
  searchWrap.innerHTML='<input class="vocab-search" id="vocabSearch" placeholder="\u641C\u7D22 / \u805A\u7126"><button class="search-clear" id="searchClear">\u2715</button>';
  controlsEl.appendChild(searchWrap);

  var catSel=document.createElement('select');catSel.className='cat-filter';catSel.id='catFilter';
  catSel.style.display=isSpecial?'none':'';
  catSel.innerHTML='<option value="">\u5168\u90E8\u5206\u7C7B</option>';
  CATEGORIES.forEach(function(c){catSel.innerHTML+='<option value="'+esc(c)+'">'+esc(c)+'</option>';});
  controlsEl.appendChild(catSel);

  var psSel=document.createElement('select');psSel.className='page-size-select';psSel.id='pageSizeSelect';
  [50,100,200,500,1000].forEach(function(n){psSel.innerHTML+='<option value="'+n+'"'+(n===100?' selected':'')+'>'+n+'\u884C</option>';});
  psSel.innerHTML+='<option value="0">\u5168\u90E8\u884C</option>';
  controlsEl.appendChild(psSel);

  var ctRow=document.getElementById('colTogglesRow');
  if(ctRow){ctRow.innerHTML='';

  var dg=document.createElement('div');dg.className='ctrl-group';
  dg.innerHTML='<span class="ctrl-group-label">\u6570\u636E\u6E90</span><button class="ctrl-chip'+(state.dataSource==='loom'?' active':'')+'" data-source="loom">Loom</button><button class="ctrl-chip'+(state.dataSource==='llpsi'?' active':'')+'" data-source="llpsi">LLPSI</button>';
  ctRow.appendChild(dg);

  if(!isSpecial){
    COL_DEFS.forEach(function(d){
      if(d.tabs.indexOf(state.tab)===-1)return;
      var b=document.createElement('button');b.className='col-toggle-btn'+(state.colVisible[d.key]?'':' off');b.textContent=d.label;
      b.onclick=function(e){e.stopPropagation();state.colVisible[d.key]=!state.colVisible[d.key];renderHeader();state.page=0;state.filteredCache=null;renderCards(true);};
      ctRow.appendChild(b);
    });
  } else {
    // Special tabs: use SPECIAL_COL_MAP
    var specCols = SPECIAL_COL_MAP[state.tab] || [];
    specCols.forEach(function(d){
      var b=document.createElement('button');b.className='col-toggle-btn'+(state.colVisible[d.key]?'':' off');b.textContent=d.label;
      b.onclick=function(e){e.stopPropagation();state.colVisible[d.key]=!state.colVisible[d.key];renderHeader();renderCards(true);};
      ctRow.appendChild(b);
    });
  }

  var count=document.createElement('span');count.className='toolbar-count';count.id='toolbarCount';
  ctRow.appendChild(count);

  dg.querySelectorAll('.ctrl-chip').forEach(function(btn){
    btn.addEventListener('click',function(ev){
      ev.stopPropagation();
      dg.querySelectorAll('.ctrl-chip').forEach(function(b){b.classList.remove('active');});
      btn.classList.add('active');
      state.dataSource=btn.dataset.source;
      resetAndRender();
    });
  });

  psSel.addEventListener('change',function(e){state.pageSize=parseInt(e.target.value)||99999;resetAndRender();});
  catSel.addEventListener('change',function(e){state.category=e.target.value;resetAndRender();});
  }

  var si=searchWrap.querySelector('#vocabSearch');
  var cl=searchWrap.querySelector('#searchClear');
  if(state.query){si.value=state.query;cl.classList.add('visible');}
  si.addEventListener('input',function(e){
    state.query=e.target.value.toLowerCase();
    cl.classList.toggle('visible',e.target.value.length>0);
    resetAndRender();
  });
  cl.addEventListener('click',function(){si.value='';state.query='';cl.classList.remove('visible');si.focus();resetAndRender();});
}

function highlightRoot(r){
  if(state.activeRootKey===r){clearRootFilter();return;}
  state.activeRootKey=r;
  updateRootBar();
  state.page=0;state.filteredCache=null;renderCards(true);
}
function clearRootFilter(){
  state.activeRootKey=null;
  updateRootBar();
  state.page=0;state.filteredCache=null;renderCards(true);
}
function updateRootBar(){
  var bar=document.getElementById('rootFilterBar');
  if(!bar)return;
  if(state.activeRootKey){
    var count=(window.rootMap[state.activeRootKey]||[]).length;
    var s3Count=(window.rootMapS3[state.activeRootKey]||[]).length;
    var info='\u8BCD\u6839: <span class="root-label">'+esc(state.activeRootKey)+'</span> \u2014 '+count+' \u4E2A\u540C\u6839\u8BCD';
    if(s3Count>0)info+=' \u00B7 '+s3Count+' \u4E2A\u5E0C\u814A\u8BCD\u6839\u5173\u8054';
    bar.innerHTML=info+' <button class="root-close" onclick="clearRootFilter()">\u2715 \u6E05\u9664</button>';
    bar.classList.add('visible');
  }else{
    bar.classList.remove('visible');
  }
}

function renderCellContent(item,col,vl){
  var inner='';
  if(col==='greek_root'){
    var gr='';
    if(item.root_match){var m=item.root_match;if(m.greek){gr='<span class="greek-col-word"><span class="word-text">'+esc(m.greek)+'</span>'+forvoBtn(m.greek,'grc')+(m.transliteration?'<span class="greek-translit">'+esc(m.transliteration)+'</span>':'')+'</span>';}else if(m.latin_root){gr='<span class="greek-col-word" style="opacity:.7"><span class="word-text">'+esc(m.latin_root)+'</span></span>';}}
    inner=gr||'<span style="color:var(--text2)">\u2014</span>';
  }else if(col==='category'){
    inner=item.category?'<span class="grid-cell-cat">'+esc(item.category)+'</span>':'';
  }else if(col==='pos'){
    var p=item.pos;var pl={noun:'\u540D\u8BCD',verb:'\u52A8\u8BCD',adjective:'\u5F62\u5BB9\u8BCD',other:'\u5176\u4ED6'};
    inner=p?'<span class="grid-cell-tag">'+(pl[p]||p)+'</span>':'';
  }else if(vl.indexOf(col)!==-1){
    var val=item[col==='latin'?'latin':col];
    var wordText=val?(typeof val==='object'?val.value||'':val):'';
    var isLat=col==='latin'&&val;
    var rm='';
    if(col==='latin'&&item.root_match&&item.root_match.latin_root){
      var rootKey=item.root_match.latin_root.toLowerCase().trim();
      var isActive=state.activeRootKey===rootKey;
      rm=' <span class="grid-cell-tag'+(isActive?' root-active':'')+'" onclick="event.stopPropagation();highlightRoot(\''+escJs(rootKey)+'\')">'+esc(item.root_match.latin_root)+'</span>';
    }
    var aiCls=(item[col+'_ai']?' ai-generated':'');
    inner='<span class="lang-word'+(isLat?' latin-root':'')+aiCls+'" onclick="fillSearch(\''+escJs(wordText)+'\')"><span class="word-text">'+esc(wordText||'\u2014')+'</span>'+(wordText?forvoBtn(wordText,col):'')+'</span>'+rm;
  }else{
    var ic=getInfoContent(item,col);
    var colDef=COL_DEFS.find(function(x){return x.key===col;});
    var colLabel=colDef?colDef.label:col;
    var storeIdx='ft'+(fullTextCounter++);
    fullTextStore[storeIdx]={full:ic,label:colLabel};
    inner='<span class="grid-cell-text" data-ft="'+storeIdx+'">'+esc(ic)+'</span>';
  }
  return inner;
}

function _headerBottom(){
  var h=document.querySelector('.header');
  return h?h.offsetHeight:0;
}
function renderCards(reset){
  var cg=document.getElementById('cardList');
  var tabRenders={s3roots:renderS3,phrases:renderPhrases,loanwords:renderLoanwords,mythology:renderMyth,bible:renderBible,encoding:renderEnc};
  if(tabRenders[state.tab]){tabRenders[state.tab](cg,reset);return;}
  if(reset){cg.innerHTML='';}
  var items=filterItems();
  var PS=state.pageSize;
  if(!items.length){cg.innerHTML='<div class="no-results">\u672A\u627E\u5230</div>';document.getElementById('loadMoreBtn').style.display='none';return;}

  var vl=getVisibleLangs(state.tab);var vi=getVisibleInfo();
  var allCols=vl.concat(vi);
  var N=allCols.length;
  var initCols=Array(N).fill('minmax(130px,1.5fr)').join(' ');

  // Sticky header wrapper (outside table-scroll, so sticky works with body scroll)
  var hdrH=_headerBottom();
  var hw=document.createElement('div');hw.className='table-header-wrap';hw.id='tableHeaderWrap';
  hw.style.top=hdrH+'px';
  var hgrid=document.createElement('div');hgrid.className='vocab-grid header-grid';
  hgrid.style.gridTemplateColumns=initCols;
  allCols.forEach(function(l){
    var d=COL_DEFS.find(function(x){return x.key===l;});
    var c=document.createElement('div');c.className='grid-header-cell';c.textContent=d?d.label:l;
    hgrid.appendChild(c);
  });
  hw.appendChild(hgrid);
  cg.appendChild(hw);

  // Scrollable data area
  var ts=document.createElement('div');ts.className='table-scroll';ts.id='dataTableScroll';
  var grid=document.createElement('div');grid.className='vocab-grid data-grid';
  grid.style.gridTemplateColumns=initCols;

  var start=state.page*PS;var end=Math.min(start+PS,items.length);
  for(var i=start;i<end;i++){
    (function(item,rowIdx){
      var key=item.english+'-'+item.category;if(grid.querySelector('[data-key="'+esc(key)+'"]'))return;
      var isRoot=false;
      if(state.activeRootKey){
        var ll=item.latin||item.lingua_latina;
        if(ll){var lv=typeof ll==='object'?ll.value:ll;if(lv&&lv.split(',')[0].trim().toLowerCase()===state.activeRootKey)isRoot=true;}
      }
      var row=document.createElement('div');row.className='grid-row';
      allCols.forEach(function(col){
        var cell=document.createElement('div');cell.className='grid-cell'+(isRoot?' root-highlight':'');
        cell.dataset.rowIdx=rowIdx;
        cell.innerHTML='<div class="grid-cell-inner">'+renderCellContent(item,col,vl)+'</div>';
        cell.addEventListener('mouseenter',function(){highlightRow(this.dataset.rowIdx,true);});
        cell.addEventListener('mouseleave',function(){highlightRow(this.dataset.rowIdx,false);});
        row.appendChild(cell);
      });
      grid.appendChild(row);
    })(items[i],i);
  }
  ts.appendChild(grid);
  cg.appendChild(ts);

  // Sync horizontal scroll between header and data
  ts.addEventListener('scroll',function(){hw.scrollLeft=ts.scrollLeft;});

  // Auto-fit columns on data grid, then sync to header grid
  requestAnimationFrame(function(){
    var colWidths=autoFitColumns(grid);
    if(colWidths){
      hgrid.style.gridTemplateColumns=colWidths.map(function(w){return w+'px'}).join(' ');
    }
    enableColResize(grid,hgrid.querySelectorAll('.grid-header-cell'),hgrid,hw,ts);
  });

  var lm=document.getElementById('loadMoreBtn');
  var tc=document.getElementById('toolbarCount');if(tc)tc.textContent='\u663E\u793A '+Math.min(end,items.length)+' / '+items.length+' \u6761';
  if(lm){lm.style.display='block';lm.disabled=end>=items.length;lm.textContent=end<items.length?'\u52A0\u8F7D\u66F4\u591A ('+end+'/'+items.length+')':'\u6CA1\u6709\u66F4\u591A\u4E86';}
  requestAnimationFrame(function(){scanTruncations(ts);});
}

function highlightRow(rowIdx,hover){
  var cells=document.querySelectorAll('.grid-cell[data-row-idx="'+rowIdx+'"]');
  cells.forEach(function(c){
    if(hover)c.classList.add('row-hover');
    else c.classList.remove('row-hover');
  });
}

function buildTable(container,reset,colLabels,items,renderCell,label,visKeys){
  if(reset){container.innerHTML='';}
  else{}
  var PS=state.pageSize;
  var pageKey=label;
  if(!state._subPages)state._subPages={};
  if(reset)state._subPages[pageKey]=0;
  var page=state._subPages[pageKey]||0;
  if(!items.length){container.innerHTML='<div class="no-results">\u672A\u627E\u5230</div>';document.getElementById('loadMoreBtn').style.display='none';return;}

  // Filter hidden columns based on colVisible
  var colIdx=[];
  var filteredLabels=[];
  colLabels.forEach(function(l,i){
    if(visKeys&&visKeys[i]&&state.colVisible[visKeys[i]]===false)return;
    colIdx.push(i);
    filteredLabels.push(l);
  });
  if(!filteredLabels.length){container.innerHTML='<div class="no-results">\u6240\u6709\u5217\u5DF2\u9690\u85CF</div>';document.getElementById('loadMoreBtn').style.display='none';return;}
  var N=filteredLabels.length;
  var initCols=Array(N).fill('minmax(130px,1.5fr)').join(' ');

  var hdrH2=_headerBottom();
  var hw=document.createElement('div');hw.className='table-header-wrap';
  hw.style.top=hdrH2+'px';
  var hgrid=document.createElement('div');hgrid.className='vocab-grid header-grid';
  hgrid.style.gridTemplateColumns=initCols;
  filteredLabels.forEach(function(l){
    var c=document.createElement('div');c.className='grid-header-cell';c.textContent=l;
    hgrid.appendChild(c);
  });
  hw.appendChild(hgrid);
  container.appendChild(hw);

  var ts=document.createElement('div');ts.className='table-scroll';
  var grid=document.createElement('div');grid.className='vocab-grid data-grid';
  grid.style.gridTemplateColumns=initCols;

  var start=page*PS;var end=Math.min(start+PS,items.length);
  items.slice(start,end).forEach(function(r,idx){
    var row=document.createElement('div');row.className='grid-row';
    var cells=renderCell(r);
    colIdx.forEach(function(ci){
      var cell=document.createElement('div');cell.className='grid-cell';
      cell.dataset.rowIdx='s'+idx;
      cell.innerHTML='<div class="grid-cell-inner">'+(cells[ci]||'')+'</div>';
      cell.addEventListener('mouseenter',function(){highlightRow(this.dataset.rowIdx,true);});
      cell.addEventListener('mouseleave',function(){highlightRow(this.dataset.rowIdx,false);});
      row.appendChild(cell);
    });
    grid.appendChild(row);
  });
  ts.appendChild(grid);
  container.appendChild(ts);

  ts.addEventListener('scroll',function(){hw.scrollLeft=ts.scrollLeft;});

  requestAnimationFrame(function(){
    var colWidths=autoFitColumns(grid);
    if(colWidths){
      hgrid.style.gridTemplateColumns=colWidths.map(function(w){return w+'px'}).join(' ');
    }
    enableColResize(grid,hgrid.querySelectorAll('.grid-header-cell'),hgrid,hw,ts);
  });

  var lm=document.getElementById('loadMoreBtn');
  var tc=document.getElementById('toolbarCount');if(tc)tc.textContent=label+' '+Math.min(end,items.length)+' / '+items.length+' \u6761';
  if(lm){lm.style.display='block';lm.disabled=end>=items.length;lm.textContent=end<items.length?'\u52A0\u8F7D\u66F4\u591A ('+end+'/'+items.length+')':'\u6CA1\u6709\u66F4\u591A\u4E86';}
  requestAnimationFrame(function(){scanTruncations(ts);});
}

function cellText(text,label){
  var storeIdx='ft'+(fullTextCounter++);
  fullTextStore[storeIdx]={full:text||'',label:label||''};
  return '<span class="grid-cell-text" data-ft="'+storeIdx+'">'+esc(text||'')+'</span>';
}

function renderS3(c,reset){
  buildTable(c,reset,['\u5E0C\u814A\u8BED','\u97F3\u8BD1','\u82F1\u6587\u91CA\u4E49','\u62C9\u4E01\u8BCD\u6839','\u793A\u4F8B','\u5206\u7EC4'],gf(SHEET3_ROOTS,['greek','transliteration','english_meaning','latin_root','english_examples','section']),function(r){
    var rootKey=r.latin_root?r.latin_root.toLowerCase().trim():'';
    var isActive=state.activeRootKey===rootKey;
    var rootTag=r.latin_root?'<span class="grid-cell-tag'+(isActive?' root-active':'')+'" onclick="event.stopPropagation();highlightRoot(\''+escJs(rootKey)+'\')">'+esc(r.latin_root)+'</span>':'';
    return ['<span class="lang-word"><span class="word-text">'+esc(r.greek)+'</span>'+forvoBtn(r.greek,'grc')+'</span>','<span style="font-style:italic;color:var(--text)">'+esc(r.transliteration)+'</span>',cellText(r.english_meaning||r.chinese_meaning,'\u82F1\u6587\u91CA\u4E49'),rootTag,cellText(r.english_examples,'\u793A\u4F8B'),'<span class="grid-cell-tag">'+esc(r.section)+'</span>'];
  },'\u8BCD\u6839\u8BCD\u7F00',['s3_greek','s3_transliteration','s3_english_meaning','s3_latin_root','s3_examples','s3_section']);
}
function renderPhrases(c,reset){
  buildTable(c,reset,['\u62C9\u4E01\u8BED','\u82F1\u8BED','\u4E2D\u6587'],gf(LATIN_PHRASES,['latin','english','chinese']),function(r){
    return ['<span class="lang-word" onclick="fillSearch(\''+escJs(r.latin)+'\')"><span class="word-text">'+esc(r.latin)+'</span>'+forvoBtn(r.latin,'lat')+'</span>',cellText(r.english,'\u82F1\u8BED'),cellText(r.chinese,'\u4E2D\u6587')];
  },'\u62C9\u4E01\u77ED\u8BED',['phr_latin','phr_english','phr_chinese']);
}
function renderLoanwords(c,reset){
  buildTable(c,reset,['\u5355\u8BCD','\u82F1\u6587\u8BF4\u660E','\u4E2D\u6587\u8BF4\u660E','\u793A\u4F8B'],gf(LOANWORDS,['word','english_note','chinese_note','example']),function(r){
    return ['<span class="lang-word" onclick="fillSearch(\''+escJs(r.word)+'\')"><span class="word-text">'+esc(r.word)+'</span>'+forvoBtn(r.word,'eng')+'</span>',cellText(r.english_note,'\u82F1\u6587\u8BF4\u660E'),cellText(r.chinese_note,'\u4E2D\u6587\u8BF4\u660E'),cellText(r.example,'\u793A\u4F8B')];
  },'\u501F\u8BCD',['lw_word','lw_english_note','lw_chinese_note','lw_example']);
}
function renderMyth(c,reset){
  buildTable(c,reset,['\u5E0C\u814A\u540D','\u62C9\u4E01\u540D','\u5E0C\u814A\u795E','\u7F57\u9A6C\u795E','\u884C\u661F','\u5907\u6CE8'],gf(MYTHOLOGY,['greek_name','latin_name','greek_deity','roman_deity','planet','notes']),function(r){
    var n=r.notes||r.words||'';
    return ['<span class="lang-word"><span class="word-text">'+esc(r.greek_name)+'</span></span>','<span class="lang-word" style="opacity:.85"><span class="word-text">'+esc(r.latin_name)+'</span></span>',cellText(r.greek_deity,'\u5E0C\u814A\u795E'),cellText(r.roman_deity,'\u7F57\u9A6C\u795E'),cellText(r.planet,'\u884C\u661F'),cellText(n,'\u5907\u6CE8')];
  },'\u795E\u8BDD',['myth_greek','myth_latin','myth_greek_deity','myth_roman_deity','myth_planet','myth_notes']);
}
function renderBible(c,reset){
  buildTable(c,reset,['\u7ECF\u6587','\u62C9\u4E01\u539F\u6587','\u4E2D\u6587\u7FFB\u8BD1','\u5173\u952E\u8BCD','\u6869\u5185\u5BB9','\u52A9\u8BB0'],gf(BIBLE,['latin','chinese','keyword','peg_content','full_mnemonic','book']),function(r){
    return ['<span class="grid-cell-ref">'+esc(r.book)+(r.verse?':'+esc(r.verse):'')+'</span>',cellText(r.latin,'\u62C9\u4E01\u539F\u6587'),cellText(r.chinese,'\u4E2D\u6587\u7FFB\u8BD1'),'<span class="grid-cell-tag">'+esc(r.keyword)+'</span>',cellText(r.peg_content,'\u6869\u5185\u5BB9'),cellText(r.full_mnemonic||r.notes||'','\u52A9\u8BB0')];
  },'\u5723\u7ECF\u52A9\u8BB0',['bib_book','bib_latin','bib_chinese','bib_keyword','bib_peg','bib_mnemonic']);
}
function renderEnc(c,reset){
  buildTable(c,reset,['\u6570\u5B57','\u57FA\u6570\u8BCD','\u5E8F\u6570\u8BCD','\u4EBA\u540D','\u52A8\u4F5C','\u7269\u54C1','\u6545\u4E8B'],gf(ENCODING,['cardinal','ordinal','number','name','action','item','story']),function(r){
    return ['<span>'+esc(r.number)+'</span>','<span class="lang-word"><span class="word-text">'+esc(r.cardinal)+'</span>'+forvoBtn(r.cardinal,'lat')+'</span>',cellText(r.ordinal,'\u5E8F\u6570\u8BCD'),cellText(r.name,'\u4EBA\u540D'),cellText(r.action,'\u52A8\u4F5C'),cellText(r.item,'\u7269\u54C1'),cellText(r.story||'','\u6545\u4E8B')];
  },'\u6570\u5B57\u7F16\u7801',['enc_number','enc_cardinal','enc_ordinal','enc_name','enc_action','enc_item','enc_story']);
}

function showPopover(el){
  if(!el)return;
  var key=el.getAttribute('data-ft');
  if(!key||!fullTextStore[key])return;
  var data=fullTextStore[key];
  var fullText=data.full||'';
  var label=data.label||'';
  if(!fullText.trim())return;
  var existing=document.querySelector('.cell-popover');
  if(existing)existing.remove();
  var rect=el.getBoundingClientRect();
  var pop=document.createElement('div');pop.className='cell-popover';
  pop.innerHTML='<button class="popover-close">\u2715</button>'+(label?'<div class="popover-label">'+esc(label)+'</div>':'')+esc(fullText);
  document.body.appendChild(pop);
  pop.querySelector('.popover-close').addEventListener('click',function(ev){ev.stopPropagation();pop.remove();});
  var pw=pop.offsetWidth;var ph=pop.offsetHeight;
  var left=rect.right+8>window.innerWidth-pw?Math.max(8,rect.left-pw-8):rect.right+8;
  var top=rect.bottom+4>window.innerHeight-ph?Math.max(8,rect.top-ph-4):rect.bottom+4;
  pop.style.left=left+'px';pop.style.top=top+'px';
  var outsideHandler=function(ev){if(!pop.contains(ev.target)&&ev.target!==el){pop.remove();document.removeEventListener('mousedown',outsideHandler);}};
  setTimeout(function(){document.addEventListener('mousedown',outsideHandler);},50);
}
function scanTruncations(container){
  if(!container)return;
  container.querySelectorAll('.grid-cell-text').forEach(function(el){
    if(el.scrollWidth>el.clientWidth+2)el.classList.add('truncated');
    else el.classList.remove('truncated');
  });
}
function loadMore(){
  var m={s3roots:'s3Page',phrases:'phrasePage',loanwords:'loanwordPage',mythology:'mythPage',bible:'biblePage',encoding:'encPage'};
  var pk=m[state.tab];
  if(pk){state[pk]++;renderCards(false);}
  else{state.page++;renderCards(false);}
  if(state._subPages&&state.tab){
    var labelMap={s3roots:'\u8BCD\u6839\u8BCD\u7F00',phrases:'\u62C9\u4E01\u77ED\u8BED',loanwords:'\u501F\u8BCD',mythology:'\u795E\u8BDD',bible:'\u5723\u7ECF\u52A9\u8BB0',encoding:'\u6570\u5B57\u7F16\u7801'};
    var lk=labelMap[state.tab];
    if(lk&&state._subPages[lk]!==undefined)state._subPages[lk]++;
  }
  requestAnimationFrame(function(){scanTruncations(document.getElementById('cardList'));});
}
function fillSearch(v){
  var si=document.getElementById('vocabSearch');
  si.value=v;state.query=v.toLowerCase();
  var cl=document.getElementById('searchClear');if(cl)cl.classList.add('visible');
  resetAndRender();
}

function autoFitColumns(grid){
  var cells=grid.querySelectorAll('.grid-cell');
  var headers=grid.querySelectorAll('.grid-header-cell');
  var N=headers.length;
  if(!N||!cells.length)return;
  grid.style.gridTemplateColumns=Array(N).fill('auto').join(' ');
  var colWidths=[];
  for(var c=0;c<N;c++){
    var maxW=0;
    var samples=[];
    for(var j=c;j<cells.length;j+=N){
      var cw=cells[j].scrollWidth;
      if(cw>0)samples.push(cw);
      if(samples.length>=30)break;
    }
    if(samples.length){
      samples.sort(function(a,b){return a-b});
      var idx=Math.min(Math.floor(samples.length*0.8),samples.length-1);
      maxW=samples[idx];
    }
    var hw=headers[c].scrollWidth;
    colWidths.push(Math.max(hw,maxW)+20);
  }
  var finalWidths=colWidths.map(function(w){return Math.max(70,Math.min(w,450))});
  grid.style.gridTemplateColumns=finalWidths.map(function(w){return w+'px'}).join(' ');
  return finalWidths;
}
function enableColResize(grid,headerCells,hgrid,hw,cg){
  var N=headerCells.length;
  if(!N)return;
  var colWidths=autoFitColumns(grid)||[];
  if(!colWidths.length){
    colWidths=[];
    for(var i=0;i<N;i++)colWidths.push(headerCells[i].offsetWidth);
  }
  headerCells.forEach(function(hc,i){
    var handle=document.createElement('div');
    handle.className='col-resize-handle';
    hc.appendChild(handle);
    var startX=0,startW=0;
    handle.addEventListener('mousedown',function(e){
      e.preventDefault();
      e.stopPropagation();
      startX=e.clientX;
      startW=colWidths[i];
      handle.classList.add('active');
      document.body.style.cursor='col-resize';
      document.body.style.userSelect='none';
      function onMove(ev){
        var dx=ev.clientX-startX;
        var newW=Math.max(60,startW+dx);
        colWidths[i]=newW;
        grid.style.gridTemplateColumns=colWidths.map(function(w){return w+'px'}).join(' ');
        if(hgrid)hgrid.style.gridTemplateColumns=colWidths.map(function(w){return w+'px'}).join(' ');
      }
      function onUp(){
        handle.classList.remove('active');
        document.body.style.cursor='';
        document.body.style.userSelect='';
        document.removeEventListener('mousemove',onMove);
        document.removeEventListener('mouseup',onUp);
      }
      document.addEventListener('mousemove',onMove);
      document.addEventListener('mouseup',onUp);
    });
  });
}

document.addEventListener('DOMContentLoaded',init);
