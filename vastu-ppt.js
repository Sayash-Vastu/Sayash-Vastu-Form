// ═══ VASTU PPT GENERATOR (template-based) ═══
// needs: <script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
// needs: template.pptx in same folder as index.html

const PPT_ESC = s => String(s==null?'':s)
  .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');


// ═══ TEMPLATES ═══ (file, and the anchor texts to replace inside it)
const PPT_TEMPLATES = {
  ckpc: { file:'template-ckpc.pptx', label:'CKPC Properties',
          cover:{client:'CKPC', prop:'HEAD OFFICE', title:'VASTU STRATEGIC', date:'Site visited on', vr:'VR/'},
          site:{prop:'Head Office, ', facing:'The site is', facingVal:'North-East Facing'},
          obs:{prop:'Head Office, ', body:'During the site visit', headRid:'rId7', photoRid:'rId3'} },
  max:  { file:'template-max.pptx', label:'MAX Healthcare',
          cover:{client:'MAX HEALTHCARE', prop:'VR Capitol', title:'VASTU STRATEGIC', date:'Recommendation on', vr:'VR/'},
          site:{prop:'Capitol Hospital', facing:'The site is', facingVal:'North-West'},
          obs:{prop:'Capitol Hospital', body:'The', headRid:null, photoRid:'rId3'} },
  sg:   { file:'template-sg.pptx', label:'Signature Global',
          cover:{client:'SIGNATURE GLOBAL', prop:'RESIDENTIAL COMMUNITY', title:'VASTU STRATEGIC', date:'Recommendation on', vr:null},
          site:{prop:null, facing:null, facingVal:null},
          obs:{prop:null, body:'The plot is', headRid:null, photoRid:'rId3'} }
};

// replace a paragraph's text, keeping its formatting
function pptSetPara(xml, anchor, newText){
  const i = xml.indexOf('<a:t>'+anchor);
  if(i<0) return xml;
  const ps = xml.lastIndexOf('<a:p>', i);
  const pe = xml.indexOf('</a:p>', i) + 6;
  const para = xml.slice(ps, pe);
  const rPr = (para.match(/<a:rPr[^>]*\/>|<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/)||[''])[0];
  const pPr = (para.match(/<a:pPr[^>]*\/>|<a:pPr[^>]*>[\s\S]*?<\/a:pPr>/)||[''])[0];
  return xml.slice(0,ps) + '<a:p>'+pPr+'<a:r>'+rPr+'<a:t>'+PPT_ESC(newText)+'</a:t></a:r></a:p>' + xml.slice(pe);
}

// wipe all paragraphs after the first inside the body text shape
function pptCollapseBody(xml, anchor){
  const i = xml.indexOf('<a:t>'+anchor);
  if(i<0) return xml;
  const bodyEnd = xml.indexOf('</p:txBody>', i);
  const firstEnd = xml.indexOf('</a:p>', i)+6;
  return xml.slice(0,firstEnd) + xml.slice(bodyEnd);
}

const EMU = 914400;
function pptTextBox(id, name, x, y, w, h, runs){
  return '<p:sp><p:nvSpPr><p:cNvPr id="'+id+'" name="'+name+'"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
    +'<p:spPr><a:xfrm><a:off x="'+Math.round(x*EMU)+'" y="'+Math.round(y*EMU)+'"/>'
    +'<a:ext cx="'+Math.round(w*EMU)+'" cy="'+Math.round(h*EMU)+'"/></a:xfrm>'
    +'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
    +'<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>'+runs+'</p:txBody></p:sp>';
}
function pptRun(text, sz, bold, color){
  return '<a:p><a:r><a:rPr lang="en-IN" sz="'+sz+'"'+(bold?' b="1"':'')
    +' dirty="0"><a:solidFill><a:srgbClr val="'+color+'"/></a:solidFill>'
    +'<a:latin typeface="+mn-lt"/></a:rPr><a:t>'+PPT_ESC(text)+'</a:t></a:r></a:p>';
}

async function generateVastuPPT(tplKey){
  const T = PPT_TEMPLATES[tplKey || (document.getElementById('ppt-tpl')||{}).value || 'ckpc'];
  const btn = document.getElementById('ppt-btn');
  const setTxt = t => { if(btn) btn.textContent = t; };
  try{
    if(typeof JSZip==='undefined'){ alert('JSZip load nahi hua'); return; }
    setTxt('⏳ Building PPT...');

    const propName = tv('f-propname') || 'Property';
    const clientNm = (tv('f-cl-name') || 'CLIENT').toUpperCase();
    const dateStr  = tv('f-date') ? new Date(tv('f-date')).toLocaleDateString('en-IN',{day:'numeric',month:'long',year:'numeric'}) : '';
    const vrNo     = tv('f-vrno') || '';
    const facing   = FD['facing'] || '—';

    // pointers that have a note or photos
    const items = Object.keys(FIELD_LABELS)
      .filter(k => (PN[k] && PN[k].trim()) || (PP[k] && PP[k].length))
      .map(k => ({ key:k, label:FIELD_LABELS[k], val:FD[k]||'', note:(PN[k]||'').trim(), photos:PP[k]||[] }));

    if(!items.length){ alert('Koi observation ya photo nahi mila.\nPointers par observation likho ya photo add karo.'); setTxt('📊 Download PPT'); return; }

    const zip = await JSZip.loadAsync(await (await fetch(T.file)).arrayBuffer());

    // ── COVER
    let s1 = await zip.file('ppt/slides/slide1.xml').async('string');
    if(T.cover.client) s1 = pptSetPara(s1,T.cover.client, clientNm);
    if(T.cover.prop) s1 = pptSetPara(s1,T.cover.prop, propName.toUpperCase());
    if(T.cover.title) s1 = pptSetPara(s1,T.cover.title,'VASTU STRATEGIC ANALYSIS');
    if(T.cover.date) s1 = pptSetPara(s1,T.cover.date,T.cover.date+' - '+dateStr);
    if(vrNo && T.cover.vr) s1 = pptSetPara(s1,T.cover.vr, vrNo);
    zip.file('ppt/slides/slide1.xml', s1);

    // ── SITE INFO
    let s2 = await zip.file('ppt/slides/slide2.xml').async('string');
    if(T.site.prop) s2 = pptSetPara(s2,T.site.prop, propName);
    if(T.site.facing) s2 = pptSetPara(s2,T.site.facing,'The site is '+facing+'.');
    if(T.site.facingVal) s2 = pptSetPara(s2,T.site.facingVal, facing);
    zip.file('ppt/slides/slide2.xml', s2);

    // ── OBSERVATION SLIDES
    const tplXml  = await zip.file('ppt/slides/slide3.xml').async('string');
    const tplRels = await zip.file('ppt/slides/_rels/slide3.xml.rels').async('string');
    let pres      = await zip.file('ppt/presentation.xml').async('string');
    let presRels  = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
    let ct        = await zip.file('[Content_Types].xml').async('string');

    const sldIds = [...pres.matchAll(/<p:sldId id="(\d+)" r:id="(rId\d+)"\/>/g)];
    const keep   = sldIds.slice(0,2).map(m=>m[0]);
    let maxId    = Math.max(...sldIds.map(m=>+m[1]));
    let maxRid   = Math.max(...[...presRels.matchAll(/Id="rId(\d+)"/g)].map(m=>+m[1]));
    const added  = [];
    let mediaN   = 900;

    for(let i=0;i<items.length;i++){
      const it = items[i], n = 10+i;
      let x = tplXml, rels = tplRels;

      if(T.obs.prop) x = pptSetPara(x,T.obs.prop, propName);
      x = pptCollapseBody(x,T.obs.body);
      x = pptSetPara(x,T.obs.body, it.note || '(observation pending)');
      x = pptSetPara(x,'0', String(i+3).padStart(2,'0').charAt(0));
      x = pptSetPara(x,'3', String(i+3).padStart(2,'0').charAt(1));

      // remove heading image (rId7) + main photo (rId3) placeholders
      if(T.obs.headRid) x = x.replace(new RegExp('<p:pic>(?:(?!</p:pic>)[\\s\\S])*?r:embed="'+T.obs.headRid+'"[\\s\\S]*?</p:pic>','g'),'');
      const hadPhoto = it.photos.length>0;
      if(hadPhoto){
        // swap main photo with first user photo
        const b64 = it.photos[0].split(',')[1];
        const mName = 'uimg'+(++mediaN)+'.jpg';
        zip.file('ppt/media/'+mName, b64, {base64:true});
        rels = rels.replace(new RegExp('(<Relationship Id="'+T.obs.photoRid+'"[^>]*Target=")[^"]*(")'), '$1../media/'+mName+'$2');
        if(!/Extension="jpg"/.test(ct)) ct = ct.replace('<Types','<Types').replace(/(<Types[^>]*>)/,'$1<Default Extension="jpg" ContentType="image/jpeg"/>');
      }else{
        x = x.replace(new RegExp('<p:pic>(?:(?!</p:pic>)[\\s\\S])*?r:embed="'+T.obs.photoRid+'"[\\s\\S]*?</p:pic>','g'),'');
      }

      // heading text in place of the removed heading image
      const head = pptTextBox(9000+i, 'Head'+i, 0.39, 1.68, 5.4, 0.5,
        pptRun(String(i+1).padStart(2,'0')+'   '+it.label.toUpperCase(), 1400, true, 'F47D35'));
      const sub  = it.val ? pptTextBox(9500+i, 'Sub'+i, 0.42, 2.18, 5.4, 0.35, pptRun(it.val, 1050, false, '4A4A4A')) : '';
      x = x.replace('</p:spTree>', head+sub+'</p:spTree>');

      zip.file('ppt/slides/slide'+n+'.xml', x);
      zip.file('ppt/slides/_rels/slide'+n+'.xml.rels', rels);

      maxId++; maxRid++;
      added.push('<p:sldId id="'+maxId+'" r:id="rId'+maxRid+'"/>');
      presRels = presRels.replace('</Relationships>',
        '<Relationship Id="rId'+maxRid+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide'+n+'.xml"/></Relationships>');
      ct = ct.replace('</Types>',
        '<Override PartName="/ppt/slides/slide'+n+'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
    }

    pres = pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst>'+keep.join('')+added.join('')+'</p:sldIdLst>');
    presRels = presRels.replace(/<Relationship[^>]*Target="slides\/slide3\.xml"[^>]*\/>/g,'');
    ct = ct.replace(/<Override[^>]*PartName="\/ppt\/slides\/slide3\.xml"[^>]*\/>/g,'');
    zip.remove('ppt/slides/slide3.xml'); zip.remove('ppt/slides/_rels/slide3.xml.rels');

    zip.file('ppt/presentation.xml', pres);
    zip.file('ppt/_rels/presentation.xml.rels', presRels);
    zip.file('[Content_Types].xml', ct);

    const blob = await zip.generateAsync({type:'blob',compression:'DEFLATE'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = (vrNo ? vrNo.replace(/\//g,'-') : 'Vastu-Report') + ' - ' + propName + '.pptx';
    a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href), 4000);
    setTxt('📊 Download PPT');
  }catch(e){
    console.error(e);
    alert('PPT banane mein dikkat: '+e.message);
    setTxt('📊 Download PPT');
  }
}
