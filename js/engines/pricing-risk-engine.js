/* Evidence-first price comparison and time-phased pricing indicators. */
(function (window) {
  'use strict';
  var NAME = 'PricingRiskEngine', VERSION = '1.0.0';
  function text(v) { return v === null || v === undefined ? '' : String(v).trim(); }
  function norm(v) { return text(v).toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, ' ').trim(); }
  function finite(v) { return typeof v === 'number' && Number.isFinite(v); }
  function pathKey(path) { return (Array.isArray(path) ? path : []).slice().sort(function (a,b) { return a.levelIndex-b.levelIndex; }).map(function (x) { return x.levelIndex + ':' + norm(x.value); }).join('|'); }
  function itemSource(item, filename, role) {
    var s=item && item.source || {}; return { workbook: filename || null, role: role, itemId: item && item.id || null,
      code: item && (item.analysisCode || item.code) || null, description: item && item.description || null,
      sheet:s.sheet || null,row:s.row || null,cells:s.cells || {},categoryPath:item && item.categoryPath || [],hierarchyStatus:item && item.hierarchyStatus || 'none' };
  }
  function specText(item) { return norm(item && (item.specification || item.spec || item.quality || item.materialSpecification)); }
  function basisIssues(bid, ref) {
    var a=bid.projectMetadata || bid.metadata || bid.project && bid.project.metadata || {}, b=ref.metadata || ref.canonical && (ref.canonical.projectMetadata || ref.canonical.metadata || ref.canonical.project && ref.canonical.project.metadata) || {}, issues=[];
    ['currency','taxIncluded','overheadPercent','profitPercent','discountPercent','baseYear','location'].forEach(function(k) {
      if (a[k] !== undefined && a[k] !== null && b[k] !== undefined && b[k] !== null && norm(a[k]) !== norm(b[k])) issues.push(k);
    });
    return { bid:a, reference:b, issues:issues, missing:['currency','taxIncluded','overheadPercent','profitPercent','discountPercent','baseYear','location'].filter(function(k){return a[k] == null || b[k] == null;}) };
  }
  function match(bid, refItems, index) {
    var bc=norm(bid.analysisCode || bid.code), bd=norm(bid.description), bu=norm(bid.unit), bs=specText(bid), candidates=[];
    if (index && bc) candidates=index.code[bc] || [];
    else if (bc) candidates=refItems.filter(function(r){return norm(r.analysisCode || r.code)===bc;});
    if (!candidates.length && bd) candidates=index ? (index.description[bd+'|'+bu] || []) : refItems.filter(function(r){return norm(r.description)===bd && bu && norm(r.unit)===bu;});
    if (!candidates.length) return {status:'NOT_FOUND', candidates:[], comparable:false, reason:'Tidak ada kandidat dengan bukti identitas memadai.'};
    if (candidates.length>1) return {status:'AMBIGUOUS',candidates:candidates,comparable:false,reason:'Lebih dari satu item pembanding cocok; tidak dipilih otomatis.'};
    var r=candidates[0], rc=norm(r.analysisCode || r.code), rd=norm(r.description), ru=norm(r.unit), rs=specText(r);
    if (bu && ru && bu!==ru) return {status:'INCOMPATIBLE_UNIT',candidates:candidates,comparable:false,reason:'Satuan berbeda dan tidak ada konversi tervalidasi.'};
    if (bs && rs && bs!==rs) return {status:'AMBIGUOUS',candidates:candidates,comparable:false,reason:'Spesifikasi berbeda.'};
    if (bs !== rs && (bs || rs)) return {status:'POSSIBLE_MATCH',candidates:candidates,comparable:false,reason:'Spesifikasi hanya tersedia pada salah satu sumber.'};
    if (bid.categoryPath && r.categoryPath && bid.categoryPath.length && r.categoryPath.length && pathKey(bid.categoryPath)!==pathKey(r.categoryPath))
      return {status:'POSSIBLE_MATCH',candidates:candidates,comparable:false,reason:'Konteks hierarki berbeda; item tidak dibandingkan lintas path tanpa verifikasi.'};
    if (bc && rc===bc && bd!==rd) return {status:'POSSIBLE_MATCH',candidates:candidates,comparable:false,reason:'Kode sama tetapi uraian berbeda; identitas pekerjaan perlu ditinjau.'};
    if (bc && rc===bc && bd===rd && bu===ru) return {status:'EXACT_MATCH',candidates:candidates,comparable:true,reason:null};
    if (!bc || !rc || bc!==rc) return {status:'POSSIBLE_MATCH',candidates:candidates,comparable:false,reason:'Pencocokan uraian saja tidak cukup untuk perbandingan otomatis.'};
    return {status:'HIGH_CONFIDENCE_MATCH',candidates:candidates,comparable:true,reason:null};
  }
  function analyzeFrontLoading(bid, references, schedule, options, comparisons, warnings) {
    if (!schedule || schedule.status!=='mapped' || !Array.isArray(schedule.periods) || !schedule.periods.length)
      return {status:'INSUFFICIENT_DATA',reason:'Jadwal terstruktur dan terpetakan belum tersedia.',periods:[],indication:null};
    var refPrice=Object.create(null), bidById=Object.create(null); comparisons.forEach(function(c){if(c.comparable && finite(c.referenceUnitPrice)) refPrice[c.itemId]=c.referenceUnitPrice;});
    bid.workItems.forEach(function(w){bidById[w.id]=w;});
    var periods=schedule.periods.map(function(p){return {label:p.label,bidValue:0,referenceValue:0,scheduledItems:0,items:[]};}), scheduledTotals=Object.create(null);
    schedule.periods.forEach(function(p,pi){(p.items||[]).forEach(function(si){
      var item=si.itemId ? bidById[si.itemId] : null;
      if (!item && si.code) { var codeMatches=bid.workItems.filter(function(w){return norm(w.analysisCode||w.code)===norm(si.code);}); if(codeMatches.length===1)item=codeMatches[0]; }
      if(!item || !finite(si.quantity) || !finite(item.unitPrice) || !finite(refPrice[item.id])) { warnings.push('Jadwal tidak lengkap/ambigu: satu atau lebih baris tidak dapat dipetakan ke pasangan harga.'); return; }
      periods[pi].bidValue += si.quantity*item.unitPrice; periods[pi].referenceValue += si.quantity*refPrice[item.id]; periods[pi].scheduledItems++;
      scheduledTotals[item.id]=(scheduledTotals[item.id]||0)+si.quantity;
      periods[pi].items.push({itemId:item.id||null,code:item.analysisCode||item.code||null,description:item.description||null,quantity:si.quantity,bidUnitPrice:item.unitPrice,referenceUnitPrice:refPrice[item.id],bidValue:si.quantity*item.unitPrice,referenceValue:si.quantity*refPrice[item.id],source:si.source||null,bidSource:itemSource(item,bid.project&&bid.project.sourceFilename,'bid')});
    });});
    var total=periods.reduce(function(s,p){return s+p.bidValue;},0), refTotal=periods.reduce(function(s,p){return s+p.referenceValue;},0);
    var tolerance=Number(options.scheduleQuantityTolerance)||0;
    var coverageMismatch=bid.workItems.some(function(item){if(!finite(refPrice[item.id]))return false;return !finite(item.volume)||scheduledTotals[item.id]===undefined||Math.abs(scheduledTotals[item.id]-item.volume)>tolerance;});
    if(!total || !refTotal || periods.some(function(p){return !p.scheduledItems;}) || coverageMismatch) return {status:'INSUFFICIENT_DATA',reason:'Nilai atau cakupan/rekonsiliasi kuantitas jadwal tidak lengkap untuk membandingkan kurva.',periods:periods,indication:null};
    var earlyCount=Math.max(1,Math.min(periods.length,Number(options.earlyPeriodCount)||1)), bidEarly=periods.slice(0,earlyCount).reduce(function(s,p){return s+p.bidValue;},0)/total, refEarly=periods.slice(0,earlyCount).reduce(function(s,p){return s+p.referenceValue;},0)/refTotal;
    var earlyShareDelta=bidEarly-refEarly, threshold=Number(options.frontLoadingEarlyShareDelta);
    periods.reduce(function(acc,p){acc.bid+=p.bidValue/total;acc.ref+=p.referenceValue/refTotal;p.cumulativeBidShare=acc.bid;p.cumulativeReferenceShare=acc.ref;return acc;},{bid:0,ref:0});
    return {status:'COMPLETE',reason:threshold>0 ? (earlyShareDelta>=threshold?'Konsentrasi nilai awal melewati ambang konfigurasi pengguna.':'Tidak melewati ambang konfigurasi pengguna.') : 'Kurva dihitung sebagai indikator; ambang belum ditetapkan.',periods:periods,earlyPeriodCount:earlyCount,bidEarlyShare:bidEarly,referenceEarlyShare:refEarly,earlyShareDelta:earlyShareDelta,threshold:threshold>0?threshold:null,indication:threshold>0?earlyShareDelta>=threshold:null};
  }
  var Engine={engineName:NAME,name:NAME,engineVersion:VERSION,analyze:function(bid,options){
    options=options||{}; var refs=Array.isArray(options.references)?options.references:[], findings=[], comps=[], warnings=[], limitations=[], comparisonsByRef=[];
    if(!bid||!Array.isArray(bid.workItems)) throw new Error('Canonical bid memerlukan workItems array.');
    var indexedRefs=refs.map(function(ref){var items=ref&&ref.canonical&&Array.isArray(ref.canonical.workItems)?ref.canonical.workItems:[], index={code:Object.create(null),description:Object.create(null)};items.forEach(function(r){var code=norm(r.analysisCode||r.code), key=norm(r.description)+'|'+norm(r.unit);if(code)(index.code[code]||(index.code[code]=[])).push(r);if(key!=='|')(index.description[key]||(index.description[key]=[])).push(r);});return {ref:ref,index:index};});
    var items=bid.workItems.filter(function(i){return i && i.rowType!=='SUBTOTAL' && i.isSubtotal!==true;}), comparableCount=0, deltaTotal=0, bidComparableValue=0, benchmarkComparableValue=0, recalculatedBidValue=0, allAmounts=true, allImpactsKnown=true, allRecalculatedKnown=true;
    items.forEach(function(item){
      var validBid=finite(item.unitPrice), primary=null, refCandidates=[];
      indexedRefs.forEach(function(indexed){var ref=indexed.ref, canon=ref && ref.canonical; if(!canon||!Array.isArray(canon.workItems)) return; var mt=match(item,canon.workItems,indexed.index);var candidate=mt.candidates.length===1?mt.candidates[0]:null;
        var basis=basisIssues(bid,ref); var status=mt.status; if(basis.issues.length) status='INCOMPARABLE_BASIS';
        var rec={referenceId:ref.id||null,referenceType:ref.type||'USER_REFERENCE',matchStatus:status,reason:basis.issues.length?'Metadata basis berbeda: '+basis.issues.join(', ')+'.':mt.reason,
          bidItemId:item.id||null,referenceItemId:candidate&&candidate.id||null,bidSource:itemSource(item,bid.project&&bid.project.sourceFilename,'bid'),referenceSource:candidate&&itemSource(candidate,canon.project&&canon.project.sourceFilename,'reference'),
          comparable:!!(mt.comparable&&!basis.issues.length&&candidate&&validBid&&finite(candidate.unitPrice)),bidUnitPrice:validBid?item.unitPrice:null,referenceUnitPrice:candidate&&finite(candidate.unitPrice)?candidate.unitPrice:null,basis:basis};
        refCandidates.push(rec); if(!primary && rec.comparable) primary=rec;
      });
      if(refCandidates.length>1 && refCandidates.filter(function(x){return x.comparable;}).length>1) { primary=null; refCandidates.forEach(function(x){x.comparable=false;x.matchStatus='AMBIGUOUS';x.reason='Beberapa sumber pembanding layak; pilih satu sumber secara eksplisit.';}); }
      var chosen=primary||refCandidates[0]||{referenceId:null,matchStatus:'NOT_FOUND',comparable:false,reason:'Sumber pembanding belum tersedia.',bidSource:itemSource(item,bid.project&&bid.project.sourceFilename,'bid')};
      chosen.candidates=refCandidates.map(function(r){return {referenceId:r.referenceId,referenceType:r.referenceType,matchStatus:r.matchStatus,reason:r.reason,referenceItemId:r.referenceItemId,referenceSource:r.referenceSource,bidSource:r.bidSource,basis:r.basis};});
      var delta=chosen.comparable?item.unitPrice-chosen.referenceUnitPrice:null, pct=chosen.comparable&&chosen.referenceUnitPrice!==0?delta/chosen.referenceUnitPrice*100:null;
      var impact=chosen.comparable&&finite(item.volume)?delta*item.volume:null;
      var comparison={itemId:item.id||null,code:item.analysisCode||item.code||null,description:item.description||null,unit:item.unit||null,categoryPath:item.categoryPath||[],matchStatus:chosen.matchStatus,comparable:!!chosen.comparable,reason:chosen.reason,
        referenceId:chosen.referenceId||null,referenceType:chosen.referenceType||null,bidUnitPrice:validBid?item.unitPrice:null,referenceUnitPrice:chosen.comparable?chosen.referenceUnitPrice:null,deltaUnitPrice:delta,deviationPercent:pct,volume:finite(item.volume)?item.volume:null,financialImpact:impact,bidAmount:finite(item.amount)?item.amount:null,bidCalculatedAmount:validBid&&finite(item.volume)?item.unitPrice*item.volume:null,referenceCalculatedAmount:chosen.comparable&&finite(item.volume)?chosen.referenceUnitPrice*item.volume:null,weightPercent:null,referenceSource:chosen.referenceSource||null,bidSource:chosen.bidSource||itemSource(item,bid.project&&bid.project.sourceFilename,'bid'),candidates:chosen.candidates||[]};
      comps.push(comparison); if(comparison.comparable){comparableCount++;if(impact!==null)deltaTotal+=impact;else allImpactsKnown=false;if(finite(item.amount))bidComparableValue+=item.amount;else allAmounts=false;if(finite(comparison.bidCalculatedAmount))recalculatedBidValue+=comparison.bidCalculatedAmount;else allRecalculatedKnown=false;if(finite(item.volume))benchmarkComparableValue+=item.volume*comparison.referenceUnitPrice;else allRecalculatedKnown=false;}
      if(comparison.matchStatus==='INCOMPARABLE_BASIS') warnings.push('Basis pembanding tidak sebanding untuk item '+(comparison.code||comparison.itemId||'tanpa kode')+'.');
      if(!validBid) warnings.push('Harga satuan penawaran tidak tersedia untuk '+(comparison.code||comparison.itemId||'item')+'.');
      if(chosen.basis && chosen.basis.missing && chosen.basis.missing.length) warnings.push('Metadata basis belum lengkap ('+chosen.basis.missing.join(', ')+') untuk pembanding '+(chosen.referenceId||'tanpa identitas')+'; nilai nominal tidak dinormalisasi.');
    });
    var bidComparableCalculated=comps.reduce(function(s,c){return s+(c.comparable&&finite(c.bidCalculatedAmount)?Math.abs(c.bidCalculatedAmount):0);},0);
    comps.forEach(function(c){if(c.comparable&&finite(c.bidCalculatedAmount)&&bidComparableCalculated!==0)c.weightPercent=Math.abs(c.bidCalculatedAmount)/bidComparableCalculated*100;});
    var threshold=Number(options.unitPriceDeviationPercent), unbalanced=[], absoluteImpactTotal=comps.reduce(function(s,c){return s+(c.comparable&&finite(c.financialImpact)?Math.abs(c.financialImpact):0);},0);
    if(threshold>0){comps.filter(function(c){return c.comparable&&c.deviationPercent!==null&&Math.abs(c.deviationPercent)>=threshold;}).forEach(function(c){var finding={id:NAME+':unit-price-deviation:'+c.itemId,ruleId:'unit_price_deviation',severity:'WARNING',interpretation:'needs_review',title:'Penyimpangan harga satuan perlu ditinjau',explanation:'Penyimpangan '+c.deviationPercent.toFixed(2)+'% terhadap pembanding terpilih; ini bukan kesimpulan ketidakpatutan.',affectedItems:[{itemId:c.itemId,code:c.code,description:c.description}],evidence:c,sourceReferences:[c.bidSource,c.referenceSource].filter(Boolean),recommendedAction:'Verifikasi spesifikasi, lokasi, tahun dasar, pajak, dan komponen harga.'};findings.push(finding);unbalanced.push({itemId:c.itemId,deviationPercent:c.deviationPercent,financialImpact:c.financialImpact,contribution:finite(c.financialImpact)&&absoluteImpactTotal?Math.abs(c.financialImpact)/absoluteImpactTotal:null,categoryPath:c.categoryPath,sourceReferences:finding.sourceReferences});});}
    var unbalancedSummary={status:threshold>0?'CONFIGURED_INDICATORS':'NOT_CONFIGURED',thresholdPercent:threshold>0?threshold:null,items:unbalanced,netComparableImpact:allImpactsKnown?deltaTotal:null,comparableItems:comparableCount};
    if(threshold<=0) limitations.push('Ambang penyimpangan tidak dikonfigurasi; tidak ada label risiko atau skor agregat yang diterbitkan.');
    refs.forEach(function(ref){var c=ref&&ref.canonical; if(c&&c.workItems) comparisonsByRef.push({id:ref.id,type:ref.type,workItems:c.workItems.length,metadata:ref.metadata||{}});});
    var sourceReferences=comps.reduce(function(a,c){return a.concat([c.bidSource,c.referenceSource].filter(Boolean));},[]).concat(options.schedule&&options.schedule.sourceReferences||[]);
    var schedule=analyzeFrontLoading(bid,refs,options.schedule,options,comps,warnings);
    if(schedule.status==='INSUFFICIENT_DATA') findings.push({id:NAME+':front-loading-insufficient-data',ruleId:'front_loading_insufficient_data',severity:'INFO',interpretation:'inconclusive',title:'Analisis distribusi waktu belum dapat disimpulkan',explanation:schedule.reason,affectedItems:[],sourceReferences:options.schedule&&options.schedule.sourceReferences||[],recommendedAction:'Sediakan jadwal dengan kuantitas tiap item per periode dan pemetaan yang jelas.'});
    else findings.push({id:NAME+':front-loading-indicator',ruleId:'front_loading_indicator',severity:schedule.indication===true?'WARNING':'INFO',interpretation:schedule.indication===true?'needs_review':'inconclusive',title:schedule.indication===true?'Indikasi konsentrasi nilai pada periode awal':'Kurva distribusi waktu untuk ditinjau',explanation:schedule.reason, evidence:{bidEarlyShare:schedule.bidEarlyShare,referenceEarlyShare:schedule.referenceEarlyShare,earlyShareDelta:schedule.earlyShareDelta,threshold:schedule.threshold,periods:schedule.periods},affectedItems:(schedule.periods||[]).reduce(function(all,p){return all.concat((p.items||[]).map(function(i){return {itemId:i.itemId,code:i.code,description:i.description};}));},[]),sourceReferences:sourceReferences.concat(options.schedule&&options.schedule.sourceReferences||[]),recommendedAction:schedule.indication===true?'Tinjau jadwal, volume per periode, basis harga, dan struktur pembayaran kontrak.':'Gunakan kurva ini sebagai indikator; tetapkan ambang pengguna bila ingin melakukan screening.'});
    return {engineName:NAME,engineVersion:VERSION,schemaVersion:'rab-guard-engine-result/1',status:items.length&&comparableCount===items.length?'complete':'partial',summary:{bidItems:items.length,referenceSources:refs.length,comparableItems:comparableCount,unmatchedOrUncertainItems:items.length-comparableCount,comparableBidAmount:allAmounts?bidComparableValue:null,comparableImportedBidAmount:allAmounts?bidComparableValue:null,comparableRecalculatedBidAmount:allRecalculatedKnown?recalculatedBidValue:null,comparableReferenceAmount:allRecalculatedKnown?benchmarkComparableValue:null,netComparableImpact:allImpactsKnown?deltaTotal:null},priceComparisons:comps,unbalancedPricingFindings:unbalancedSummary,frontLoadingFindings:schedule,matchingResults:comps.map(function(c){return {itemId:c.itemId,status:c.matchStatus,reason:c.reason,candidates:c.candidates};}),riskIndicators:findings,findings:findings,sourceReferences:sourceReferences,assumptions:['Hanya item non-subtotal yang dianalisis.','Tidak ada konversi satuan atau normalisasi pajak/biaya yang diasumsikan.','Ambang hanya digunakan jika diberikan secara eksplisit oleh pengguna.'],limitations:limitations,warnings:warnings,errors:[],metrics:{matching:'kode/uraian exact dan satuan; ambiguous candidate tidak dibandingkan',deviationFormula:'(bidUnitPrice-referenceUnitPrice)/referenceUnitPrice*100',amountImpactFormula:'(bidUnitPrice-referenceUnitPrice)*volume',riskScore:'not provided'},referenceSources:comparisonsByRef};
  }};
  window.PricingRiskEngine=Engine;
})(typeof window!=='undefined'?window:this);
