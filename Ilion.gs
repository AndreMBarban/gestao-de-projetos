/***** CONFIGURACAO *****/
// 01_Responsaveis
const RESP_FOLDER_ID = '1EbXzzyS9tmTdy4NDcYPB14cNpjjS0D3Y';
// 02_Projetos
const PROJ_FOLDER_ID = '166KSPir-g-r5XDFJ9s_rQ-NWnxK8HWW8';
// (opcional) Template Google Docs para primeira pauta de um projeto novo (pode deixar vazio '')
const TEMPLATE_DOC_ID = '';
const TIMEZONE       = 'America/Sao_Paulo';

/***** HEADINGS / TIPOS *****/
const H1 = DocumentApp.ParagraphHeading.HEADING1;
const H2 = DocumentApp.ParagraphHeading.HEADING2;
const NORMAL = DocumentApp.ParagraphHeading.NORMAL;

/***** REGEX / MARCADORES *****/
const MENTION_RE = /@([A-Za-zÀ-ÿ0-9][\wÀ-ÿ'’.\- ]]+)/g;
// Prefixo visível de índice na ata por responsável: "7.1.3 – "
const INDEX_SEP = ' – ';
const INDEX_PREFIX_RE = /^\s*(\d+(?:\.\d+)+)\s*[–-]\s+/;

/***** FUZZY (fallback) *****/
const FUZZY_MIN_SCORE = 0.70;
const FUZZY_MARGIN    = 0.10;
const DIAG_FUZZY_LOG  = false;

/***** UTILITARIOS *****/
function yymmdd_(d) { d = d || new Date(); return Utilities.formatDate(d, TIMEZONE, 'yyMMdd'); }
function escapeRegExp_(s){ return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function nbspToSpace_(s){ return String(s||'').replace(/\u00A0/g,' '); }
function collapseSpaces_(s){ return String(s||'').replace(/\s+/g,' ').trim(); }
function normalizeTxt_(s){ return collapseSpaces_(nbspToSpace_(s)); }
function stripMentions_(txt){ return normalizeTxt_((txt || '').replace(MENTION_RE, '')); }
function extractMentions_(txt){
  var found = [], m;
  var base = String(txt||'');
  while ((m = MENTION_RE.exec(base)) !== null){
    var nome = collapseSpaces_(m[1].replace(/[.,;:!?)”»]+$/,''));
    if (nome && !found.some(x=>x.toLowerCase()===nome.toLowerCase())) found.push(nome);
  }
  return found;
}
function pushUnique_(arr, item, seenSet, key){
  if (seenSet.has(key)) return;
  seenSet.add(key);
  arr.push(item);
}

/***** TOKENIZAÇÃO & SIMILARIDADE (p/ fuzzy) *****/
function tokenize_(s){
  s = normalizeTxt_(s || '').toLowerCase();
  s = s.replace(/[“”"«»'’]/g,'').replace(/[.,;:!?(){}\[\]]/g,' ');
  var parts = s.split(/\s+/).filter(Boolean);
  var stop = new Set(['de','da','do','das','dos','e','&','vs','a','o','as','os','para','por','em','no','na','nos','nas','com','sem','um','uma','ao','à','às','ou']);
  return parts.filter(t=>!stop.has(t));
}
function bigrams_(s){
  s = normalizeTxt_(s || '').toLowerCase().replace(/\s+/g,' ');
  var out = [];
  for (var i=0;i<s.length-1;i++){ out.push(s.substring(i,i+2)); }
  return out;
}
function jaccard_(a, b){
  var A = new Set(a), B = new Set(b);
  if (A.size===0 && B.size===0) return 1;
  var inter = 0; A.forEach(x=>{ if (B.has(x)) inter++; });
  var uni = A.size + B.size - inter;
  return uni ? inter/uni : 0;
}
function dice_(a, b){
  if (!a.length && !b.length) return 1;
  var A = new Map(); a.forEach(x=>A.set(x,(A.get(x)||0)+1));
  var inter = 0;
  b.forEach(x=>{ var c=A.get(x)||0; if (c>0){ inter++; A.set(x,c-1); } });
  return (2*inter)/(a.length + b.length || 1);
}
function comboSim_(s1, s2){ return 0.6*jaccard_(tokenize_(s1), tokenize_(s2)) + 0.4*dice_(bigrams_(s1), bigrams_(s2)); }

/***** ESTILOS *****/
function setParaStyle_(p, size, bold){ if (!p) return; try{p.setBold(bold===true);}catch(e){} try{p.setFontSize(size);}catch(e){} }
function styleListItem_(li){
  if (!li) return;
  try { li.setGlyphType(DocumentApp.GlyphType.BULLET); } catch(e){}
  try { li.setBold(false); } catch(e){}
  try { li.setFontSize(10); } catch(e){}
}
function forceAllBullets_(doc){
  var body = doc.getBody();
  for (var i=0;i<body.getNumChildren();i++){
    var el = body.getChild(i);
    if (el.getType()===DocumentApp.ElementType.LIST_ITEM){
      var li = el.asListItem();
      try{li.setGlyphType(DocumentApp.GlyphType.BULLET);}catch(e){}
      try{li.setBold(false);}catch(e){}
      try{li.setFontSize(10);}catch(e){}
    }
  }
}

/***** NUMERAÇÃO DE CAPÍTULOS (H1) *****/
function stripLeadingNumber_(s){ return String(s||'').replace(/^\s*\d+\s*[\.\)\-–]\s+/,'').trim(); }
function renumberChapters_(doc){
  var body = doc.getBody(), idx=0;
  for (var i=0;i<body.getNumChildren();i++){
    var el = body.getChild(i);
    if (el.getType()!==DocumentApp.ElementType.PARAGRAPH) continue;
    var p = el.asParagraph(); if (p.getHeading()!==H1) continue;
    var raw = p.getText()||''; var title = stripLeadingNumber_(raw);
    idx++; p.setText(idx+'. '+title); try{p.setBold(true);}catch(e){}
  }
}

/***** DRIVE HELPERS *****/
function listDocsInFolder_(folderId){
  var it=DriveApp.getFolderById(folderId).getFilesByType(MimeType.GOOGLE_DOCS), out=[];
  while (it.hasNext()) out.push(it.next());
  return out;
}
function createEmptyDocInFolder_(name, folderId){
  var doc=DocumentApp.create(name);
  var file=DriveApp.getFileById(doc.getId());
  DriveApp.getFolderById(folderId).addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return doc;
}
function makeCopyInFolder_(fileId, name, folderId){
  var file=DriveApp.getFileById(fileId).makeCopy(name, DriveApp.getFolderById(folderId));
  return DocumentApp.openById(file.getId());
}
function replaceEverywhere_(doc, pattern, repl){
  doc.getBody().replaceText(pattern,repl);
  var h=doc.getHeader(); if(h)h.replaceText(pattern,repl);
  var f=doc.getFooter(); if(f)f.replaceText(pattern,repl);
}

/***** NOMES DE PROJETO / ARQUIVOS *****/
function listProjectDocs_(){
  var re=/^\s*\d{6}\s*[–-]?\s*Pauta\s+Projeto\s*[–-]?\s+/i;
  return listDocsInFolder_(PROJ_FOLDER_ID).filter(f=>re.test(nbspToSpace_(f.getName())));
}
function parseProjectNameFromFilename_(name){
  var n=normalizeTxt_(name);
  var m=n.match(/^\s*(\d{6})\s*(?:[–-]\s*)?Pauta\s+Projeto\s*(?:[–-]\s*)?(.*)$/i);
  return m?collapseSpaces_(m[2]):n;
}
function getProjectDocByStamp_(projectName, stamp){
  var target=stamp+' Pauta Projeto '+projectName;
  var folder=DriveApp.getFolderById(PROJ_FOLDER_ID);
  var it=folder.getFilesByName(target);
  return it.hasNext()?DocumentApp.openById(it.next().getId()):null;
}
function getLatestProjectDoc_(projectName){
  var pname=escapeRegExp_(normalizeTxt_(projectName));
  var re=new RegExp('^(\\d{6})\\s*[–-]?\\s*Pauta\\s+Projeto\\s*[–-]?\\s*'+pname+'$','i');
  var items=listDocsInFolder_(PROJ_FOLDER_ID).filter(f=>re.test(normalizeTxt_(f.getName())));
  var best=null,bestStamp=-1;
  items.forEach(f=>{
    var m=normalizeTxt_(f.getName()).match(/^(\d{6})\s*[–-]?\s*Pauta/i);
    var stamp=m?parseInt(m[1],10):-1; if(stamp>bestStamp){bestStamp=stamp; best=f;}
  });
  return best?DocumentApp.openById(best.getId()):null;
}

/***** CAPÍTULOS (H1) *****/
function getChaptersFromProjectDoc_(doc){
  var body=doc.getBody(), chapters=[], seen=new Set();
  for (var i=0;i<body.getNumChildren();i++){
    var el=body.getChild(i);
    if (el.getType()!==DocumentApp.ElementType.PARAGRAPH) continue;
    var p=el.asParagraph();
    if (p.getHeading()===H1){
      var t=normalizeTxt_(p.getText());
      if (t && !seen.has(t.toLowerCase())) { chapters.push(t); seen.add(t.toLowerCase()); }
    }
  }
  return chapters;
}

/***** DETECÇÃO/CONVERSÃO DE CAPÍTULOS EM LISTA NIVEL 0 -> H1 *****/
function hasAnyH1_(doc){
  var body=doc.getBody();
  for (var i=0;i<body.getNumChildren();i++){
    var el=body.getChild(i);
    if (el.getType()===DocumentApp.ElementType.PARAGRAPH && el.asParagraph().getHeading()===H1) return true;
  }
  return false;
}
function convertListChaptersToH1_(doc){
  var body=doc.getBody();
  if (hasAnyH1_(doc)) return false;
  var targets=[];
  for (var i=0;i<body.getNumChildren();i++){
    var el=body.getChild(i);
    if (el.getType()===DocumentApp.ElementType.LIST_ITEM){
      var li=el.asListItem(); if (li.getNestingLevel()===0) targets.push(i);
    }
  }
  if (!targets.length) return false;
  for (var k=targets.length-1;k>=0;k--){
    var idx=targets[k]; var li=body.getChild(idx).asListItem();
    var txt=normalizeTxt_(li.getText())||'Capítulo';
    var p=body.insertParagraph(idx, txt); p.setHeading(H1); safeRemove_(li);
  }
  return true;
}

/***** REMOCAO SEGURA *****/
function isParaLike_(el){ if(!el)return false; var t=el.getType(); return t===DocumentApp.ElementType.PARAGRAPH||t===DocumentApp.ElementType.LIST_ITEM; }
function countParaLikeInBody_(body){ var n=0; for (var i=0;i<body.getNumChildren();i++){ var ch=body.getChild(i); if(isParaLike_(ch)) n++; } return n; }
function safeRemove_(el){
  if (!el) return; var parent=el.getParent(); if(!parent) return;
  if (parent.getType()===DocumentApp.ElementType.BODY_SECTION){
    var body=parent; var total=countParaLikeInBody_(body);
    if (isParaLike_(el) && total<=1){
      try{ if (el.getType()===DocumentApp.ElementType.LIST_ITEM) el.asListItem().setText(''); else el.asParagraph().setText(''); }catch(e){} return;
    }
  }
  try{ el.removeFromParent(); }
  catch(e){ try{ if (el.getType()===DocumentApp.ElementType.LIST_ITEM) el.asListItem().setText(''); else el.asParagraph().setText(''); }catch(_){ } }
}

/** Limpa conteúdo abaixo de capítulos (preservando a “casca” H1). */
function cleanBulletsUnderChapters_(doc){
  var body=doc.getBody(), toRemove=[], placeholder=body.appendParagraph('\u200B');
  if (hasAnyH1_(doc)){
    var inside=false;
    for (var i=0;i<body.getNumChildren();i++){
      var el=body.getChild(i), type=el.getType();
      if (type===DocumentApp.ElementType.PARAGRAPH){
        var p=el.asParagraph(); if (p.getHeading()===H1){ inside=true; continue; }
        if (inside) toRemove.push(p);
      } else if (type===DocumentApp.ElementType.LIST_ITEM){
        if (inside) toRemove.push(el.asListItem());
      }
    }
  } else {
    var seen0=false;
    for (var j=0;j<body.getNumChildren();j++){
      var el2=body.getChild(j), t2=el2.getType();
      if (t2===DocumentApp.ElementType.LIST_ITEM){
        var li=el2.asListItem(), lvl=li.getNestingLevel();
        if (lvl===0){ seen0=true; continue; }
        if (seen0 && lvl>=1) toRemove.push(li);
      } else if (t2===DocumentApp.ElementType.PARAGRAPH){
        var p2=el2.asParagraph(); if (seen0) toRemove.push(p2);
      }
    }
  }
  for (var k=toRemove.length-1;k>=0;k--) safeRemove_(toRemove[k]);
  try{
    var parent=placeholder.getParent();
    if (parent && parent.getType()===DocumentApp.ElementType.BODY_SECTION){
      var total=countParaLikeInBody_(parent);
      if (total>1) placeholder.removeFromParent(); else placeholder.setText('');
    } else { placeholder.setText(''); }
  }catch(e){ try{ placeholder.setText(''); }catch(_){ } }
}

/***** SUPORTE A CHIPS DE PESSOA (Google Docs) *****/
function extractPeopleChipsFromContainer_(container){
  var names=[]; function add_(s){ s=collapseSpaces_(String(s||'').replace(/^<|>$/g,'')); if(!s)return; var k=s.toLowerCase(); if(!names.some(x=>x.toLowerCase()===k)) names.push(s); }
  if (!container || typeof container.getNumChildren!=='function') return names;
  for (var i=0;i<container.getNumChildren();i++){
    var child=container.getChild(i), t=child.getType();
    if (DocumentApp.ElementType.PERSON && t===DocumentApp.ElementType.PERSON){ try{ var person=child.asPerson(); add_(person.getName()||person.getEmail()||''); continue; }catch(e){} }
    if (t===DocumentApp.ElementType.TEXT){
      var textEl=child.asText(), txt=textEl.getText(), idxs=textEl.getTextAttributeIndices();
      for (var j=0;j<idxs.length;j++){
        var start=idxs[j], end=(j+1<idxs.length)?idxs[j+1]-1:txt.length-1;
        var url=textEl.getLinkUrl(start);
        if (url && (/^mailto:/i.test(url)||/people|contact|profile|google\.com\/u\//i.test(url))){ add_(txt.substring(start,end+1)); }
      }
    }
  }
  return names;
}

/***** PARSE PROJETO — cap -> [ { text, level, chips } ] *****/
function parseProjectDoc_(doc){
  var out={}, body=doc.getBody(), cap=null, seenAnyH1=hasAnyH1_(doc);
  function ensure_(c){ if(!out[c]) out[c]=[]; }
  function push_(c,text,level,container){ if(!c)return; var txt=normalizeTxt_(text||''); if(!txt)return; var chips=extractPeopleChipsFromContainer_(container); ensure_(c); out[c].push({text:txt, level:level||0, chips:chips}); }
  for (var i=0;i<body.getNumChildren();i++){
    var el=body.getChild(i), type=el.getType();
    if (type===DocumentApp.ElementType.PARAGRAPH){
      var p=el.asParagraph(), txt=normalizeTxt_(p.getText());
      if (p.getHeading()===H1){ var title=txt; if (!title) continue; cap=title; ensure_(cap); continue; }
      if (cap && p.getHeading()!==H1 && txt){ push_(cap, txt, 0, p); }
      continue;
    }
    if (type===DocumentApp.ElementType.LIST_ITEM){
      var li=el.asListItem(), liText=normalizeTxt_(li.getText()), lvl=li.getNestingLevel();
      if (seenAnyH1){ if (liText) push_(cap, liText, lvl, li); }
      else {
        if (lvl===0){ cap=liText||'Capítulo'; ensure_(cap); }
        else { var normalized=Math.max(0,lvl-1); if (liText) push_(cap, liText, normalized, li); }
      }
      continue;
    }
  }
  return out;
}

/***** DETECTOR DE CAPA DE ATA POR RESPONSÁVEL (H1 que devemos ignorar) *****/
function isResponsibleCoverTitle_(txt){
  var t=normalizeTxt_(txt||''); if(!t) return false;
  if (/^tarefas\s*[–-]?\s*de\b/i.test(t)) return true;
  if (/semana\s*\d{6}/i.test(t)) return true;
  return false;
}

/***** Nome do responsável (derivado do nome do arquivo) *****/
function parseResponsibleNameFromDocName_(name){
  var n=normalizeTxt_(name||'');
  var m=n.match(/^tarefas\s*[–-]\s*(.+?)\s*[–-]\s*semana/i);
  return m ? collapseSpaces_(m[1]) : collapseSpaces_(n.replace(/^tarefas\s*[–-]\s*/i,''));
}

/***** PARSE RESPONSÁVEL – guarda idxPath (ex.: [7,1,1]) *****/
function parseResponsibleDoc_(doc, opts){
  var author = opts && opts.author;
  var out={}, body=doc.getBody(), proj=null, cap=null, globalIdx=0;
  function extractIndex_(t){
    var m=(t||'').match(INDEX_PREFIX_RE);
    if (!m) return { clean: normalizeTxt_(t), path: null };
    var parts = m[1].split('.').map(n=>parseInt(n,10));
    var clean = normalizeTxt_(t.replace(INDEX_PREFIX_RE,''));
    return { clean: clean, path: parts };
  }
  function push_(t,lvl){
    if (!proj || !cap) return;
    var ext = extractIndex_(t);
    out[proj]=out[proj]||{}; out[proj][cap]=out[proj][cap]||[];
    out[proj][cap].push({ text: ext.clean, level: lvl||0, _orig: globalIdx++, idx: ext.path, author: author });
  }
  for (var i=0;i<body.getNumChildren();i++){
    var el=body.getChild(i);
    if (el.getType()===DocumentApp.ElementType.PARAGRAPH){
      var p=el.asParagraph(), txt=normalizeTxt_(p.getText());
      if (p.getHeading()===H1){ if (isResponsibleCoverTitle_(txt)){ proj=null; cap=null; continue; } proj=normalizeTxt_(txt); cap=null; if (proj && !out[proj]) out[proj]={}; continue; }
      if (p.getHeading()===H2){ cap=normalizeTxt_(txt); if (proj && cap && !out[proj][cap]) out[proj][cap]=[]; continue; }
      if (!proj||!cap) continue;
      if (p.getHeading()!==H1 && txt){ push_(txt, 0); }
    } else if (el.getType()===DocumentApp.ElementType.LIST_ITEM){
      if (!proj||!cap) continue;
      var li=el.asListItem(), txt2=normalizeTxt_(li.getText()), lvl2=li.getNestingLevel();
      if (txt2){ push_(txt2, lvl2); }
    }
  }
  return out;
}

/***** Criar/atualizar pauta do projeto *****/
function upsertProjectDoc_(projectName, stampYYMMDD){
  var doc=getProjectDocByStamp_(projectName, stampYYMMDD);
  if (!doc){
    var base=getLatestProjectDoc_(projectName);
    if (base){ doc=makeCopyInFolder_(base.getId(), stampYYMMDD+' Pauta Projeto '+projectName, PROJ_FOLDER_ID); }
    else { if (TEMPLATE_DOC_ID) doc=makeCopyInFolder_(TEMPLATE_DOC_ID, stampYYMMDD+' Pauta Projeto '+projectName, PROJ_FOLDER_ID);
           else doc=createEmptyDocInFolder_(stampYYMMDD+' Pauta Projeto '+projectName, PROJ_FOLDER_ID); }
  }
  convertListChaptersToH1_(doc);
  cleanBulletsUnderChapters_(doc);
  replaceEverywhere_(doc, '{{\\s*PROJETO\\s*}}', projectName);
  doc.saveAndClose();
  return DocumentApp.openById(doc.getId());
}

/***** Inserção sob capítulo preservando níveis *****/
function insertUnderChapter_(doc, chapter, items){
  if (!items||!items.length) return;
  var body=doc.getBody(), chNorm=normalizeTxt_(chapter), idxH1=-1;
  for (var i=0;i<body.getNumChildren();i++){
    var el=body.getChild(i);
    if (el.getType()!==DocumentApp.ElementType.PARAGRAPH) continue;
    var p=el.asParagraph();
    if (p.getHeading()===H1 && normalizeTxt_(p.getText()).toLowerCase()===chNorm.toLowerCase()){ idxH1=i; break; }
  }
  if (idxH1===-1){ var ph=body.appendParagraph(chapter); ph.setHeading(H1); body.appendParagraph(''); idxH1=body.getChildIndex(ph); }
  var insertAt=idxH1+1;
  body.insertParagraph(insertAt++,'');
  items.forEach(function(it){
    var li=body.insertListItem(insertAt++, it.text);
    try{ if (it.level && it.level>0) li.setNestingLevel(it.level); }catch(e){}
    styleListItem_(li);
  });
}

/***** Índices da pauta-base (por capítulo) *****/
// info = {
//   order[cap][pathKey] = seq,
//   parentMax[cap][parentKey] = lastSeqAmongDesc,
//   levelIndex[cap][level] = [ { key, text, seq, lastSeq } ... ],
//   chapterIndex[cap] = 1..N,
//   indexOrder[cap][idxStr] = { seq, lastSeq }   // "7.1.2" → onde está
// }
function buildBaseOrderInfoForProject_(projectName){
  var base=getLatestProjectDoc_(projectName);
  var info={ order:{}, parentMax:{}, levelIndex:{}, chapterIndex:{}, indexOrder:{} };
  if (!base) return info;

  var chapters = getChaptersFromProjectDoc_(base); // ordem visual
  chapters.forEach(function(c, idx){ info.chapterIndex[c]=idx+1; });

  var byCap=parseProjectDoc_(base);
  Object.keys(byCap).forEach(function(cap){
    var seq=0;
    info.order[cap]=info.order[cap]||{};
    info.parentMax[cap]=info.parentMax[cap]||{};
    info.levelIndex[cap]=info.levelIndex[cap]||{};
    info.indexOrder[cap]=info.indexOrder[cap]||{};

    var stackTxt=[], levelCounters=[];

    byCap[cap].forEach(function(it){
      stackTxt.splice(it.level);
      stackTxt[it.level]=stripMentions_(it.text);

      if (levelCounters[it.level]==null) levelCounters[it.level]=0;
      levelCounters[it.level] += 1;
      for (var d=it.level+1; d<levelCounters.length; d++) levelCounters[d]=0;

      var key = stackTxt.slice(0, it.level+1).join(' > ') + ' @L' + it.level;
      var idxNumbers = [ info.chapterIndex[cap] ].concat(levelCounters.slice(0, it.level+1).map(x=>x||0));
      var idxStr = idxNumbers.join('.');

      seq += 1;
      info.order[cap][key]=seq;

      if (!info.levelIndex[cap][it.level]) info.levelIndex[cap][it.level]=[];
      info.levelIndex[cap][it.level].push({ key:key, text:stackTxt[it.level], seq:seq, lastSeq:seq, idxStr:idxStr });

      info.indexOrder[cap][idxStr] = info.indexOrder[cap][idxStr] || { seq: seq, lastSeq: seq };

      for (var lvl=0; lvl<=it.level-1; lvl++){
        var pkey = stackTxt.slice(0,lvl+1).join(' > ') + ' @L' + lvl;
        var cur = info.parentMax[cap][pkey] || 0;
        if (seq>cur) info.parentMax[cap][pkey]=seq;
      }
    });

    // segundo passe: lastSeq por nó e propagação ao pai
    var lvls = Object.keys(info.levelIndex[cap]).map(x=>parseInt(x,10)).sort((a,b)=>b-a);
    var lastByKey={};
    lvls.forEach(function(lvl){
      (info.levelIndex[cap][lvl]||[]).forEach(function(node){
        var last = lastByKey[node.key] || node.seq;
        node.lastSeq = Math.max(node.seq, last);
        lastByKey[node.key] = node.lastSeq;

        if (lvl>0){
          var parts = node.key.split(' > '); parts.pop();
          var parentKey = parts.join(' > ') + ' @L' + (lvl-1);
          lastByKey[parentKey] = Math.max(lastByKey[parentKey]||0, node.lastSeq);
        }
        var io = info.indexOrder[cap][node.idxStr];
        if (io) io.lastSeq = Math.max(io.lastSeq||node.seq, node.lastSeq);
      });
    });
  });
  return info;
}

/***** Fuzzy parent no mesmo cap/nível-1 (fallback) *****/
function fuzzyParentKey_(baseInfo, chapter, level, parentText){
  if (!baseInfo || !baseInfo.levelIndex || !baseInfo.levelIndex[chapter]) return null;
  var candLevel = level - 1; if (candLevel<0) return null;
  var cands = baseInfo.levelIndex[chapter][candLevel] || []; if (!cands.length) return null;
  var scored = cands.map(c=>({ key:c.key, text:c.text, seq:c.seq, lastSeq:c.lastSeq, score:comboSim_(parentText,c.text) }))
                    .sort((a,b)=>b.score-a.score);
  if (!scored.length) return null;
  var best=scored[0], second=scored[1]||null;
  var margin = second ? (best.score-second.score) : best.score;
  if (DIAG_FUZZY_LOG){
    Logger.log('[FUZZY] chap="%s" lvl=%s parent="%s" -> best="%s" score=%.3f margin=%.3f', chapter, candLevel, parentText, best.text, best.score, margin);
  }
  if (best.score>=FUZZY_MIN_SCORE && margin>=FUZZY_MARGIN) return best;
  return null;
}

/***** Consolidação: Responsáveis -> Projetos (ancoragem por índice; fallback fuzzy) *****/
function consolidateFromPeopleToProjects(stampYYMMDDOpt){
  var agg={}, respFiles=listDocsInFolder_(RESP_FOLDER_ID), orderCache={}, parentNewCounters={}, topNewCounters={};

  function mergeIntoBucket_(bucket, dedupKey, incoming){
    if (!bucket.map) bucket.map={};
    var existing=bucket.map[dedupKey];
    if (!existing){
      incoming.authors = incoming.author ? [incoming.author] : [];
      incoming.obs = [];
      incoming.idxStr = incoming.idxStr || null;
      bucket.map[dedupKey]=incoming;
      bucket.items.push(incoming);
      return;
    }

    existing.ord = Math.min(existing.ord, incoming.ord);
    existing._orig = Math.min(existing._orig, incoming._orig);

    if (!existing.idxStr && incoming.idxStr) existing.idxStr = incoming.idxStr;

    existing.authors = existing.authors || [];
    if (incoming.author && existing.authors.every(a=>a.toLowerCase()!==incoming.author.toLowerCase())){
      existing.authors.push(incoming.author);
    }

    if (incoming.author){
      existing.obs = existing.obs || [];
      var noteExists = existing.obs.some(o=>o.author===incoming.author && o.text===incoming.text);
      var shouldAddNote = (!noteExists && incoming.text);
      if (shouldAddNote){ existing.obs.push({ author: incoming.author, text: incoming.text }); }
    }
  }

  function renderTextWithObs_(item){
    if (!item) return '';

    var allAuthors = (item.authors||[]).map(a=>String(a||''));
    var notes = (item.obs||[]).slice();

    // garante que todo autor apareça pelo menos uma vez (mesmo sem divergência de texto)
    allAuthors.forEach(function(a){
      var hasNote = notes.some(n=>String(n.author||'').toLowerCase()===a.toLowerCase());
      if (!hasNote) notes.push({ author: a, text: null });
    });

    if (!notes.length) return item.text;

    var obsTxt = notes.map(function(o){
      var author=collapseSpaces_(o.author||'Autor');
      var extra=(o.text && o.text!==item.text) ? (': '+o.text) : '';
      return 'Obs — '+author+extra;
    }).join('; ');

    return item.text + ' (' + obsTxt + ')';
  }

  function nextAfterParentSubtree_(baseInfo, cap, parentIdxStr){
    // coloca novo filho após o último seq do pai
    var io = (baseInfo.indexOrder[cap]||{})[parentIdxStr];
    var last = io ? (io.lastSeq||io.seq||0) : 0;
    if (!parentNewCounters[cap]) parentNewCounters[cap]={};
    parentNewCounters[cap][parentIdxStr] = (parentNewCounters[cap][parentIdxStr]||0) + 1;
    return last + parentNewCounters[cap][parentIdxStr];
  }
  function nextTopNew_(proj, cap){
    if (!topNewCounters[proj]) topNewCounters[proj]={};
    var n=(topNewCounters[proj][cap]||0)+1; topNewCounters[proj][cap]=n;
    return 1000000000 + n;
  }
  function bucketHasItems_(bucket){ return !!(bucket && bucket.items && bucket.items.length); }
  function projectHasAnyItems_(proj){ if(!agg[proj])return false; return Object.keys(agg[proj]).some(c=>bucketHasItems_(agg[proj][c])); }

  respFiles.forEach(function(f){
    var doc=DocumentApp.openById(f.getId());
    var respName=parseResponsibleNameFromDocName_(f.getName());
    var mapa=parseResponsibleDoc_(doc, { author: respName });

    Object.keys(mapa).forEach(function(proj){
      var P=normalizeTxt_(proj); if(!agg[P]) agg[P]={};
      if (!orderCache[P]) orderCache[P]=buildBaseOrderInfoForProject_(P);
      var baseInfo=orderCache[P];

      Object.keys(mapa[proj]).forEach(function(cap){
        var C=normalizeTxt_(cap); if(!agg[P][C]) agg[P][C]={ items:[], seen:new Set(), map:{} };

        var stack=[];
        var stackClean=[];
        (mapa[proj][cap]||[]).forEach(function(it){
          stack.splice(it.level); stackClean.splice(it.level);
          stack[it.level]=it.text; stackClean[it.level]=stripMentions_(it.text);

          // 1) tentativa por índice numérico (ex.: [7,1,1])
          var ord = null;
          if (it.idx && it.idx.length >= 2) { // [capIdx, bulletIdx, ...]
            var idxStr = it.idx.join('.');
            var io = (baseInfo.indexOrder[C] || {})[idxStr];

            // chave textual deste item (para saber se já existia no projeto-base)
            var keyByText = stackClean.slice(0, it.level + 1).join(' > ') + ' @L' + it.level;
            var existsInBaseByText = ((baseInfo.order[C] || {})[keyByText] != null);

            if (io && io.seq != null && existsInBaseByText) {
              // item que já existia no projeto-base -> mantém posição exata
              ord = io.seq;
            } else if (it.idx.length > 2) {
              // novo sub-bullet → ancora após a subárvore do PAI (índice sem o último segmento)
              var parentIdxStr = it.idx.slice(0, it.idx.length - 1).join('.');
              ord = nextAfterParentSubtree_(baseInfo, C, parentIdxStr);
            } else {
              // novo nível 0 → fim do capítulo
              ord = nextTopNew_(P, C);
            }
          }

          // 2) se não deu por índice, tenta pelo caminho de texto (se existir no base)
          if (ord==null){
            var key = stackClean.slice(0, it.level+1).join(' > ') + ' @L' + it.level;
            var ordBase = ((baseInfo.order[C]||{})[key]!=null) ? baseInfo.order[C][key] : null;
            if (ordBase!=null){ ord = ordBase; }
          }

          // 3) se ainda não, tenta fuzzy pro pai
          if (ord==null){
            if (it.level>0){
              var parentText = stackClean[it.level-1] || '';
              var fuzzy = fuzzyParentKey_(baseInfo, C, it.level, parentText);
              if (fuzzy && fuzzy.key){
                var parentNode = (baseInfo.levelIndex[C] && baseInfo.levelIndex[C][it.level-1] || []).find(n=>n.key===fuzzy.key);
                var parentIdxStr = parentNode ? parentNode.idxStr : null;
                if (parentIdxStr) ord = nextAfterParentSubtree_(baseInfo, C, parentIdxStr);
                else ord = (fuzzy.lastSeq||0) + 1;
              } else {
                ord = nextTopNew_(P, C);
              }
            } else {
              ord = nextTopNew_(P, C);
            }
          }

          var dedupKey = (it.idx && it.idx.length) ? ('IDX:'+it.idx.join('.')+'@L'+it.level) : (stackClean.slice(0, it.level+1).join(' > ') + ' @L' + it.level);
          mergeIntoBucket_(agg[P][C], dedupKey, { text: stackClean[it.level], level: it.level, ord: ord, _orig: it._orig||0, idxStr: it.idx ? it.idx.join('.') : null, author: it.author });
        });
      });
    });
  });

  // Ordena e grava
  Object.keys(agg).forEach(function(proj){
    Object.keys(agg[proj]).forEach(function(cap){
      agg[proj][cap].items.sort(function(a,b){ if(a.ord!==b.ord) return a.ord-b.ord; if(a.level!==b.level) return a.level-b.level; return a._orig-b._orig; });
      agg[proj][cap].items = agg[proj][cap].items.map(function(it){ return { text: renderTextWithObs_(it), level: it.level }; });
    });
  });

  var stamp = stampYYMMDDOpt || yymmdd_();
  Object.keys(agg).forEach(function(project){
    if (!projectHasAnyItems_(project)) return;
    var doc=upsertProjectDoc_(project, stamp);
    var chapters=getChaptersFromProjectDoc_(doc);
    chapters.forEach(function(ch){ var bucket=agg[project][ch]; if (bucket && bucket.items.length) insertUnderChapter_(doc, ch, bucket.items); });
    Object.keys(agg[project]).forEach(function(ch){ var exists=chapters.some(c=>c.toLowerCase()===ch.toLowerCase()); if(!exists) insertUnderChapter_(doc, ch, agg[project][ch].items); });
    renumberChapters_(doc); forceAllBullets_(doc); doc.saveAndClose();
  });
}

/***** Split: Projetos -> Responsáveis (agora com prefixo numérico visível) *****/
function splitProjectsToPeopleFromProjects(stampYYMMDDOpt){
  var projFiles=listProjectDocs_(); if(!projFiles.length) throw new Error('Nenhum doc de Projeto encontrado em 02_Projetos.');
  var mapaResp={};

  projFiles.forEach(function(f){
    var project=parseProjectNameFromFilename_(f.getName());
    var doc=DocumentApp.openById(f.getId());
    var byCap=parseProjectDoc_(doc);
    var chapters=getChaptersFromProjectDoc_(doc); // ordem
    var chapterIdxMap={}; chapters.forEach(function(c,idx){ chapterIdxMap[c]=idx+1; });

    Object.keys(byCap).forEach(function(cap){
      var items=byCap[cap], stack=[], stackClean=[], levelCounters=[];
      items.forEach(function(it){
        stack.splice(it.level); stackClean.splice(it.level);
        stack[it.level]=it.text; stackClean[it.level]=stripMentions_(it.text);

        // atribuição/menções
        var assignees=[].concat(it.chips||[], extractMentions_(it.text)||[]);
        assignees=assignees.filter((v,i,a)=>a.findIndex(x=>String(x).toLowerCase()===String(v).toLowerCase())===i);
        if (!assignees.length) return;

        // calcula prefixo numérico (capIdx.levelIdx[.sub…])
        var capIdx = chapterIdxMap[cap] || 0;
        if (levelCounters[it.level]==null) levelCounters[it.level]=0;
        levelCounters[it.level] += 1;
        for (var d=it.level+1; d<levelCounters.length; d++) levelCounters[d]=0;

        var idxParts=[capIdx].concat(levelCounters.slice(0, it.level+1).map(x=>x||0));
        var idxStr=idxParts.join('.');

        var cleanText=stripMentions_(it.text);
        var textWithIndex = idxStr + INDEX_SEP + cleanText;

        assignees.forEach(function(r){
          var R=normalizeTxt_(r);
          if(!mapaResp[R]) mapaResp[R]={};
          if(!mapaResp[R][project]) mapaResp[R][project]={};
          if(!mapaResp[R][project][cap]) mapaResp[R][project][cap]={ items:[], seen:new Set() };

          // pais (contexto)
          for (var lvl=0; lvl<it.level; lvl++){
            var parentTxt = stackClean[lvl]; if (!parentTxt) continue;
            var parentIdxParts=[capIdx].concat(levelCounters.slice(0,lvl+1).map(x=>x||0));
            var parentIdxStr=parentIdxParts.join('.');
            var parentWithIndex = parentIdxStr + INDEX_SEP + parentTxt;

            var keyP = stackClean.slice(0, lvl+1).join(' > ') + ' @L' + lvl;
            pushUnique_(mapaResp[R][project][cap].items, { text: parentWithIndex, level:lvl }, mapaResp[R][project][cap].seen, keyP);
          }

          // o próprio item
          var myKey = stackClean.slice(0, it.level).concat([cleanText]).join(' > ') + ' @L' + it.level;
          pushUnique_(mapaResp[R][project][cap].items, { text: textWithIndex, level: it.level }, mapaResp[R][project][cap].seen, myKey);
        });
      });
    });
  });

  var stamp=stampYYMMDDOpt||yymmdd_();
  var folder=DriveApp.getFolderById(RESP_FOLDER_ID);

  Object.keys(mapaResp).forEach(function(r){
    var name='Tarefas – '+r+' – Semana '+stamp;
    var file, it=folder.getFilesByName(name);
    if (it.hasNext()){ file=it.next(); DocumentApp.openById(file.getId()).getBody().clear(); }
    else { var d=DocumentApp.create(name); var f=DriveApp.getFileById(d.getId()); folder.addFile(f); DriveApp.getRootFolder().removeFile(f); file=f; }

    var doc=DocumentApp.openById(file.getId()), body=doc.getBody();
    var t1=body.appendParagraph('Tarefas de '+r+' — Semana '+stamp); t1.setHeading(H1); setParaStyle_(t1,12,true);

    Object.keys(mapaResp[r]).sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'})).forEach(function(proj){
      var ph1=body.appendParagraph(proj); ph1.setHeading(H1); setParaStyle_(ph1,12,true);

      // capítulos na ordem do projeto-base, se houver
      var baseDoc=getLatestProjectDoc_(proj);
      var orderedChs=baseDoc?getChaptersFromProjectDoc_(baseDoc):Object.keys(mapaResp[r][proj]);

      orderedChs.forEach(function(ch){
        var bucket=mapaResp[r][proj][ch]; if (!bucket||!bucket.items.length) return;
        var ph2=body.appendParagraph(ch); ph2.setHeading(H2); setParaStyle_(ph2,10,true);
        bucket.items.forEach(function(itm){
          var li=body.appendListItem(itm.text);
          try{ if (itm.level && itm.level>0) li.setNestingLevel(itm.level); }catch(e){}
          styleListItem_(li);
        });
        body.appendParagraph('');
      });

      Object.keys(mapaResp[r][proj]).forEach(function(ch){
        if (orderedChs.some(c=>c.toLowerCase()===ch.toLowerCase())) return;
        var ph2b=body.appendParagraph(ch); ph2b.setHeading(H2); setParaStyle_(ph2b,10,true);
        (mapaResp[r][proj][ch].items||[]).forEach(function(itm){
          var li2=body.appendListItem(itm.text);
          try{ if (itm.level && itm.level>0) li2.setNestingLevel(itm.level); }catch(e){}
          styleListItem_(li2);
        });
        body.appendParagraph('');
      });

      body.appendParagraph('');
    });

    forceAllBullets_(doc);
    doc.saveAndClose();
  });
}

/***** MENU & GATILHOS *****/
function createTriggers() {
  deleteTriggers_();
  ScriptApp.newTrigger('consolidateFromPeopleToProjects')
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(19).inTimezone(TIMEZONE).create();
  // ScriptApp.newTrigger('splitProjectsToPeopleFromProjects')
  //   .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY)
  //   .atHour(21).inTimezone(TIMEZONE).create();
}
function deleteTriggers_(){ ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t)); }

/***** DIAGNÓSTICOS *****/
function _diagnoseAnchorsExample(){
  var info=buildBaseOrderInfoForProject_('Exemplo Projeto');
  Logger.log('%s', JSON.stringify(info.indexOrder||{}, null, 2));
}
function _diagnoseSplitChips(){
  Logger.log('Pasta 02: %s', PROJ_FOLDER_ID);
  var files=listProjectDocs_();
  if (!files.length){ Logger.log('>>> Nenhum doc no padrão "AAMMDD Pauta Projeto <Nome>"'); return; }
  files.forEach(function(f){
    var raw=f.getName(), proj=parseProjectNameFromFilename_(raw);
    Logger.log('— %s  => Projeto: [%s]', raw, proj);
    var doc=DocumentApp.openById(f.getId()), byCap=parseProjectDoc_(doc), caps=Object.keys(byCap);
    if (!caps.length){ Logger.log('   (Sem capítulos detectados)'); return; }
    caps.forEach(function(cap){
      Logger.log('   CAP: %s', cap);
      (byCap[cap]||[]).forEach(function(it, idx){
        var chips=(it.chips||[]).join(', ')||'—', mentions=extractMentions_(it.text).join(', ')||'—';
        Logger.log('     • #%d L%d: "%s"   chips=[%s]   @=[%s]', idx+1, it.level||0, it.text, chips, mentions);
      });
    });
  });
}
