import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

const PROXY_URL = 'https://twcg-proxy.kyle-88e.workers.dev';
const SP_LIB    = 'TWCG-BD-Platform';

export default class TwcgbdPlatformWebPart extends BaseClientSideWebPart<{}> {

  private V: any = {};
  private C: any = {};
  private M: any = {};
  private PW: any = {};
  private logs: any[] = [];
  private comps: any[] = [];
  private nav: number = 0;
  private aV: string|null = null;
  private aC: string|null = null;
  private uFile: File|null = null;
  private dtCb: any = null;

  private BT = ['Small Business','WOSB','EDWOSB','VOSB','SDVOSB','8(a)','HUBZone','SDB','Emerging Large','Large Business','Commercial Vendor','Mentor','Protege','Joint Venture'];
  private CE = ['ISO 9001:2015','ISO 27001:2022','CAGE Registered','SAM Active'];
  private DC = ['Accounting System','Purchasing System','Property Mgmt','Estimating System'];

  public render(): void {
    this.domElement.innerHTML = this._html();
    this._boot();
  }

  // ── SHAREPOINT REST API ──────────────────────────────────────────────────────
  private get _spHeaders() {
    return {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'odata-version': ''
    };
  }

  private async _spGet(folder: string, file: string): Promise<any> {
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${SP_LIB}/${folder}')/Files('${file}')/$value`;
      const r = await this.context.spHttpClient.get(url, (await import('@microsoft/sp-http')).SPHttpClient.configurations.v1);
      if (!r.ok) return null;
      return await r.json();
    } catch(e) { return null; }
  }

  private async _spList(folder: string): Promise<any[]> {
    try {
      const { SPHttpClient } = await import('@microsoft/sp-http');
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${SP_LIB}/${folder}')/Files?$select=Name`;
      const r = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!r.ok) return [];
      const d = await r.json();
      return (d.value || []).filter((f: any) => f.Name.endsWith('.json'));
    } catch(e) { return []; }
  }

  private async _spSave(folder: string, file: string, data: any): Promise<void> {
    try {
      const { SPHttpClient, SPHttpClientResponse } = await import('@microsoft/sp-http');
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${SP_LIB}/${folder}')/Files/Add(url='${file}',overwrite=true)`;
      await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: this._spHeaders,
        body: JSON.stringify(data)
      });
    } catch(e) { console.error('SP save error:', e); }
  }

  private async _loadSP(): Promise<void> {
    this._showBanner('Loading from SharePoint...', 'wait');
    try {
      const cf = await this._spList('profiles-company');
      for (const f of cf) { const d = await this._spGet('profiles-company', f.Name); if (d && d.id) this.C[d.id] = d; }
      const vf = await this._spList('profiles-vehicles');
      for (const f of vf) { const d = await this._spGet('profiles-vehicles', f.Name); if (d && d.id) this.V[d.id] = d; }
      const sf = await this._spList('comparisons');
      this.comps = [];
      for (const f of sf) {
        if (f.Name === 'matrix-state.json') { const ms = await this._spGet('comparisons', f.Name); if (ms) { if (ms.M) Object.assign(this.M, ms.M); if (ms.PW) Object.assign(this.PW, ms.PW); } }
        else { const s = await this._spGet('comparisons', f.Name); if (s) this.comps.push(s); }
      }
      const lg = await this._spGet('intel-log', 'intel-log.json');
      if (lg && Array.isArray(lg)) this.logs = lg;
      this._ensureMx();
      this._renderSidebar();
      this._renderContent();
      this._showBanner('SharePoint sync complete ✓', 'ok');
    } catch(e: any) { this._showBanner('Load error: ' + e.message, 'err'); }
  }

  private _saveCo(co: any) { this._spSave('profiles-company', co.id + '.json', co); }
  private _saveV(v: any)   { this._spSave('profiles-vehicles', v.id + '.json', v); }
  private _saveMx()        { this._spSave('comparisons', 'matrix-state.json', { M: this.M, PW: this.PW, updated: new Date().toISOString() }); }
  private _saveLog()       { this._spSave('intel-log', 'intel-log.json', this.logs); }
  private _saveCompSP(c: any){ this._spSave('comparisons', c.id + '.json', c); }

  private _addLog(scope: string, sid: string, entry: string, auth?: string) {
    this.logs.push({ date: new Date().toLocaleString(), author: auth || this.context.pageContext.user.displayName, scope, sid, entry });
    this._saveLog();
  }

  // ── AI API ───────────────────────────────────────────────────────────────────
  private _callAPI(prompt: string, cb: (err: any, data: any) => void): void {
    fetch(PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ model: 'claude-sonnet-4-20250514', max_tokens: 4000, messages: [{ role: 'user', content: prompt }] })
    })
    .then(r => r.json())
    .then(d => {
      if (d.error) { cb(new Error(d.error.message), null); return; }
      const raw = (d.content && d.content[0] ? d.content[0].text : '').replace(/```json/g,'').replace(/```/g,'').trim();
      try { cb(null, JSON.parse(raw)); } catch(e) { cb(e, null); }
    })
    .catch(e => cb(e, null));
  }

  // ── HELPERS ──────────────────────────────────────────────────────────────────
  private _sc(p: number) { return p >= 55 ? '#107c10' : p >= 40 ? '#835c00' : '#a4262c'; }
  private _sbg(s: string) { const m: any = { Tracking:'#f3f2f1',Shaping:'#deecf9',Bid:'#fff4ce','No-Bid':'#fde7e9',Award:'#dff6dd',Loss:'#f3f2f1' }; return m[s]||'#f3f2f1'; }
  private _sfg(s: string) { const m: any = { Tracking:'#605e5c',Shaping:'#004578',Bid:'#835c00','No-Bid':'#a4262c',Award:'#107c10',Loss:'#a4262c' }; return m[s]||'#605e5c'; }
  private _dl(d: string)  { return d==='pursue'?'Pursue':d==='conditional'?'Conditional':'No-Go'; }
  private _dp(d: string)  { return d==='pursue'?'go':d==='conditional'?'vf':'ac'; }
  private _uid()          { return 'x' + Date.now() + Math.random().toString(36).slice(2,5); }

  private _showBanner(msg: string, type: string) {
    const b = this.domElement.querySelector('#spbanner') as HTMLElement;
    const s = this.domElement.querySelector('#spmsg') as HTMLElement;
    if (!b || !s) return;
    b.style.display = 'flex';
    b.style.background = type==='ok'?'#dff6dd':type==='err'?'#fde7e9':'#fff4ce';
    b.style.color = type==='ok'?'#107c10':type==='err'?'#a4262c':'#835c00';
    s.textContent = msg;
    if (type==='ok') setTimeout(()=>{ b.style.display='none'; }, 3000);
  }

  private _ps(id: string, t: string, m: string) {
    const e = this.domElement.querySelector('#'+id) as HTMLElement;
    if (!e) return;
    e.className = 'pstat on ' + (t==='ok'?'pok':t==='err'?'per':'pwt');
    e.textContent = m;
  }

  // ── MATRIX LOGIC ─────────────────────────────────────────────────────────────
  private _ensureMx() {
    Object.keys(this.V).forEach(vid => {
      if (!this.M[vid]) this.M[vid] = {};
      Object.keys(this.C).forEach(cid => {
        if (!this.M[vid][cid]) this.M[vid][cid] = { pursue: 'conditional', notes: '', stage: 'Tracking', po: null, so: null };
        if (!this.PW[vid]) this.PW[vid] = {};
        if (!this.PW[vid][cid]) {
          this.PW[vid][cid] = {};
          const fs = this.V[vid].parsed ? this.V[vid].parsed.pwinFactors || [] : [];
          fs.forEach((f: any) => { this.PW[vid][cid][f.key] = f.val; });
        }
      });
    });
  }

  private _cpwin(vid: string, cid: string): number {
    const p = this.V[vid]?.parsed; if (!p) return 0;
    const s = this.PW[vid]?.[cid] || {};
    const fs = p.pwinFactors || []; if (!fs.length) return 0;
    const tot = fs.reduce((a: number, f: any) => a + (s[f.key] || 1), 0);
    return Math.round((tot / (fs.length * 3)) * 100);
  }

  private _cscore(vid: string, cid: string): number|null {
    const p = this.V[vid]?.parsed; if (!p?.scoreCalcRules?.length) return null;
    const co = this.C[cid]; if (!co) return null;
    let score = 0;
    p.scoreCalcRules.forEach((rule: any) => {
      let v = '';
      if (rule.criterion==='facility_clearance') v = co.clearance||'None';
      else if (rule.criterion==='cmmc') v = co.cmmc||'0';
      else if (rule.criterion==='iso_9001') v = String(co.certs?.includes('ISO 9001:2015'));
      else if (rule.criterion==='dcma_accounting') v = String(co.dcma?.includes('Accounting System'));
      else if (rule.criterion==='business_size') v = co.size||'';
      if (rule.values?.[v] !== undefined) score += rule.values[v];
    });
    return score;
  }

  // ── PROMPTS ───────────────────────────────────────────────────────────────────
  private RFPP = 'You are Deuce, Growth Architect for TWCG. Analyze this RFP and return ONLY valid JSON with no markdown.\n{"vehicleName":"short name","agency":"issuing agency","type":"MA-IDIQ","pop":"period of performance","naics":["541611"],"ceiling":"TBD","releaseDate":"TBD","dueDate":"TBD","awardDate":"TBD","solNum":"TBD","sbSlots":"SB slot description","hasScorecard":false,"scorecardSections":[],"gates":[{"num":"01","label":"Gate label","color":"#1b3a6b","rows":[["Criterion","Requirement","VERIFY","vf"]]}],"domains":[{"name":"Domain","naics":"541611","pct":70,"rec":true,"color":"#107c10","note":"brief note","caps":[["Capability","aligned"]]}],"pwinFactors":[{"key":"rel","label":"Customer Intimacy","desc":"Agency relationships","low":"No relationships","mid":"Indirect access","high":"Decision-maker access","color":"#1b3a6b","val":1},{"key":"pp","label":"Past Performance","desc":"Relevant contracts","low":"No relevant work","mid":"Related experience","high":"Direct mission match","color":"#107c10","val":2},{"key":"comp","label":"Competitive Position","desc":"vs. field","low":"Undifferentiated","mid":"Some differentiators","high":"Clear advantage","color":"#5c2d91","val":2},{"key":"sol","label":"Solution Maturity","desc":"Readiness","low":"Early concept","mid":"Partial capability","high":"Proven and deployed","color":"#835c00","val":2},{"key":"team","label":"Teaming Strength","desc":"Partners","low":"No partners","mid":"Potential partners","high":"Established team","color":"#d83b01","val":1}],"scoreCalcRules":[{"criterion":"facility_clearance","values":{"Secret":2500,"Top Secret":3500},"label":"Clearance"},{"criterion":"cmmc","values":{"2":2500,"2c":3500,"3":3500},"label":"CMMC"},{"criterion":"iso_9001","values":{"true":2500},"label":"ISO 9001"},{"criterion":"dcma_accounting","values":{"true":3000},"label":"DCMA Accounting"},{"criterion":"business_size","values":{"Small Business":2000},"label":"SB Non-JV"}],"scoreRef":[["Category","Max Points","Key Levers"]],"defaultDecision":"conditional","decisionRationale":"2-3 sentence rationale"}\n\nRFP TEXT:\n';
  private COP  = 'You are Deuce, Growth Architect for TWCG. Extract company intelligence and return ONLY valid JSON with no markdown.\n{"name":"Full legal name","short":"Short name","cage":"CAGE or empty","uei":"UEI or empty","icon":"single emoji","hq":"City State or empty","web":"website or empty","size":"Small Business","bizTypes":["WOSB"],"naics":["541611"],"psc":["R408"],"clearance":"None","certs":["ISO 9001:2015"],"cmmc":"0","dcma":["Accounting System"],"caps":"comma separated capabilities","rev":"revenue or empty","notes":"2-3 sentence BD summary","sourceNote":"confidence note"}\n\nTWCG context: Small federal BD firm. DOD focus: Army, INDOPACOM. Partners: HII, SMX, SAIC, Booz Allen.\n\nSOURCE:\n';

  // ── HTML TEMPLATE ─────────────────────────────────────────────────────────────
  private _html(): string {
    return `
<style>
*{box-sizing:border-box;margin:0;padding:0}
.twcg{font-family:Segoe UI,sans-serif;background:#f3f2f1;color:#323130;font-size:14px;display:flex;flex-direction:column;height:800px;overflow:hidden;width:100%}
.topbar{background:#1b3a6b;color:#fff;padding:9px 20px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.badge{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);border-radius:20px;padding:3px 10px;font-size:11px;font-weight:600}
.nav{background:#fff;border-bottom:1px solid #edebe9;display:flex;flex-shrink:0;padding:0 16px;overflow-x:auto}
.ntab{padding:11px 14px;font-size:12px;font-weight:500;cursor:pointer;color:#605e5c;border-bottom:3px solid transparent;white-space:nowrap;background:none;border-left:none;border-right:none;border-top:none;font-family:Segoe UI,sans-serif}
.ntab.on{color:#1b3a6b;border-bottom-color:#1b3a6b;font-weight:600}
.spbanner{display:none;align-items:center;gap:10px;padding:6px 16px;font-size:11px;font-weight:600;flex-shrink:0}
.body{display:flex;flex:1;overflow:hidden}
.sidebar{width:230px;min-width:230px;background:#fff;border-right:1px solid #edebe9;display:flex;flex-direction:column;overflow:hidden}
.content{flex:1;overflow-y:auto;padding:14px}
.sbs{border-bottom:1px solid #edebe9;padding:10px;flex-shrink:0}
.sbt{font-size:10px;font-weight:700;color:#605e5c;text-transform:uppercase;letter-spacing:.06em;margin-bottom:7px}
.uzone{border:2px dashed #c8c6c4;border-radius:6px;padding:10px;text-align:center;cursor:pointer;background:#faf9f8}
.pastebtn,.addbtn{width:100%;padding:7px;border:1px solid #edebe9;border-radius:5px;background:#faf9f8;font-size:11px;color:#605e5c;cursor:pointer;font-family:Segoe UI,sans-serif;margin-top:5px}
.addbtn{background:#1b3a6b;color:#fff;font-weight:600;border:none}
.sitem{padding:7px 9px;border-radius:6px;cursor:pointer;margin-bottom:3px;border:1px solid transparent;display:flex;align-items:center;gap:7px}
.sitem:hover{background:#f3f2f1}.sitem.on{background:#e8edf5;border-color:#1b3a6b}
.sico{width:26px;height:26px;border-radius:5px;display:flex;align-items:center;justify-content:center;font-size:11px;flex-shrink:0;font-weight:700}
.sinf{flex:1;min-width:0}.snm{font-size:12px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.smt{font-size:10px;color:#605e5c}
.stag{font-size:9px;font-weight:700;padding:1px 5px;border-radius:8px}.sok{background:#dff6dd;color:#107c10}.swt{background:#fff4ce;color:#835c00}.srw{background:#f3f2f1;color:#605e5c}
.dx{font-size:11px;color:#a19f9d;cursor:pointer;padding:2px 4px;opacity:0}.sitem:hover .dx{opacity:1}
.card{background:#fff;border-radius:8px;border:1px solid #edebe9;margin-bottom:10px;overflow:hidden}
table.t{width:100%;border-collapse:collapse;font-size:12px}
table.t th{text-align:left;font-size:10px;font-weight:600;color:#605e5c;padding:5px 7px;border-bottom:1px solid #edebe9;background:#faf9f8}
table.t td{padding:6px 7px;border-bottom:1px solid #f3f2f1;vertical-align:top}
table.t tr:last-child td{border-bottom:none}
.mxw{overflow-x:auto;border-radius:8px;border:1px solid #edebe9}
.mx{border-collapse:collapse;width:100%;font-size:11px}
.mx th{background:#1b3a6b;color:#fff;padding:8px 10px;text-align:center;font-weight:600;white-space:nowrap;min-width:120px}
.mx th:first-child{text-align:left;min-width:140px}
.mx td{padding:8px 10px;border:1px solid #edebe9;text-align:center;vertical-align:middle;cursor:pointer}
.mx td:first-child{text-align:left;font-weight:600;background:#faf9f8;cursor:default}
.ci{display:flex;flex-direction:column;align-items:center;gap:2px}.cpw{font-size:15px;font-weight:700}
.cbar{width:44px;height:3px;background:#edebe9;border-radius:2px;overflow:hidden}.cbarf{height:3px;border-radius:2px}
.pill{display:inline-block;font-size:10px;font-weight:700;padding:2px 7px;border-radius:10px}
.go{background:#dff6dd;color:#107c10}.vf{background:#fff4ce;color:#835c00}.ac{background:#fde7e9;color:#a4262c}.sc{background:#deecf9;color:#004578}.as{background:#f3f2f1;color:#323130}
.spill{font-size:9px;font-weight:700;padding:1px 5px;border-radius:8px;margin-top:1px}
.overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;align-items:center;justify-content:center}
.overlay.on{display:flex}
.modal{background:#fff;border-radius:10px;padding:20px;width:92%;max-width:580px;max-height:90vh;overflow-y:auto}
.modal h2{font-size:15px;font-weight:700;margin-bottom:14px;color:#1b3a6b}
.fr{margin-bottom:10px}.fr label{display:block;font-size:11px;font-weight:600;color:#605e5c;margin-bottom:3px}
.fr input,.fr select,.fr textarea{width:100%;padding:7px 9px;border:1px solid #edebe9;border-radius:5px;font-size:12px;font-family:Segoe UI,sans-serif;background:#faf9f8}
.f2{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.ckg{display:flex;flex-wrap:wrap;gap:5px;margin-top:4px}
.cki{display:flex;align-items:center;gap:4px;padding:3px 8px;border:1px solid #edebe9;border-radius:12px;background:#f3f2f1;cursor:pointer;font-size:11px;user-select:none}
.cki.on{background:#e8edf5;border-color:#1b3a6b;color:#1b3a6b;font-weight:600}
.mbtns{display:flex;gap:7px;justify-content:flex-end;margin-top:14px;flex-wrap:wrap}
.btn{padding:7px 14px;border-radius:5px;font-size:12px;font-weight:600;cursor:pointer;font-family:Segoe UI,sans-serif;border:none}
.bp{background:#1b3a6b;color:#fff}.bg{background:#fff;color:#1b3a6b;border:1px solid #1b3a6b}.bc{background:#f3f2f1;color:#323130;border:1px solid #edebe9}.bd{background:#fde7e9;color:#a4262c}
.hint{font-size:11px;color:#605e5c;margin-bottom:8px;line-height:1.5;padding:8px;background:#f3f2f1;border-radius:5px}
.pstat{padding:10px;border-radius:6px;font-size:12px;margin-top:8px;display:none}.pstat.on{display:block}
.pok{background:#dff6dd;color:#107c10}.per{background:#fde7e9;color:#a4262c}.pwt{background:#fff4ce;color:#835c00}
.itbar{display:flex;border-bottom:2px solid #edebe9;margin-bottom:14px}
.itab{flex:1;padding:9px 6px;text-align:center;font-size:12px;font-weight:600;cursor:pointer;color:#605e5c;background:transparent;border:none;font-family:Segoe UI,sans-serif}
.itab.on{color:#fff;background:#1b3a6b}
.ipane{display:none}.ipane.on{display:block}
.empty{text-align:center;padding:50px 20px;color:#605e5c}
.empty h2{font-size:17px;font-weight:600;color:#1b3a6b;margin-bottom:7px}
.stitle{font-size:10px;font-weight:700;color:#605e5c;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px}
.hrow{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:7px}
.htitle{font-size:15px;font-weight:700;color:#1b3a6b}
.sgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:8px;margin-bottom:12px}
.sc2{background:#fff;border:1px solid #edebe9;border-radius:7px;padding:10px 12px}
.sl{font-size:10px;color:#605e5c;margin-bottom:3px}.sv{font-size:13px;font-weight:600}
.dtabs{display:flex;background:#fff;border-radius:7px;margin-bottom:10px;overflow:hidden;border:1px solid #edebe9;flex-wrap:wrap}
.dtab{flex:1;padding:9px 5px;text-align:center;font-size:11px;font-weight:500;cursor:pointer;color:#605e5c;border-bottom:3px solid transparent;white-space:nowrap;min-width:80px;background:none;border-left:none;border-right:none;border-top:none;font-family:Segoe UI,sans-serif}
.dtab.on{color:#1b3a6b;border-bottom-color:#1b3a6b}
.dpane{display:none}.dpane.on{display:block}
.dg{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:8px}
.dc{background:#fff;border:1px solid #edebe9;border-radius:7px;overflow:hidden}
.dt{padding:10px 12px 7px}.dn{font-size:12px;font-weight:600}
.bb{background:#edebe9;border-radius:3px;height:4px;margin:6px 0 3px}.bf{height:4px;border-radius:3px}
.cr{display:flex;justify-content:space-between;padding:4px 12px;border-top:1px solid #f3f2f1;font-size:11px}
.ali{color:#107c10;font-weight:600}.par{color:#835c00;font-weight:600}.gap{color:#a4262c;font-weight:600}
.ovr{display:flex;align-items:center;gap:8px;margin-top:8px;padding:8px;background:#f3f2f1;border-radius:6px;flex-wrap:wrap}
.ovr label{font-size:11px;font-weight:600;color:#605e5c;white-space:nowrap}
.ovr input[type=number]{width:64px;padding:4px 6px;border:1px solid #edebe9;border-radius:4px;font-size:13px;font-weight:700;text-align:center}
.slrow{display:flex;align-items:flex-start;gap:8px;margin-bottom:8px;flex-wrap:wrap}
.sli{min-width:120px}.sll{font-size:11px;font-weight:600}.sld{font-size:10px;color:#605e5c}
.opts{display:flex;gap:3px;flex:1;min-width:150px}
.opt{flex:1;padding:4px 2px;font-size:10px;border:1px solid #edebe9;border-radius:4px;background:#f3f2f1;cursor:pointer;text-align:center;font-family:Segoe UI,sans-serif;color:#323130}
.opt.on{color:#fff;border-color:transparent}
.log{padding:8px 10px;border-left:3px solid #edebe9;margin-bottom:6px;background:#faf9f8;border-radius:0 4px 4px 0;font-size:12px}
.logm{font-size:10px;color:#605e5c;margin-bottom:3px}
.tag{font-size:10px;padding:2px 6px;border-radius:8px;font-weight:600;display:inline-block;margin:2px}
</style>

<div class="twcg">
  <div class="topbar">
    <div>
      <div style="font-size:15px;font-weight:700">TWCG BD Intelligence Platform</div>
      <div style="font-size:11px;opacity:.7;margin-top:1px">Contract Vehicle - Company - Pursuit Analysis</div>
    </div>
    <div style="display:flex;align-items:center;gap:10px">
      <div class="badge">AI Agent (Deuce) [Growth Architect]</div>
      <div id="upill" style="background:rgba(255,255,255,.2);border-radius:14px;padding:3px 10px;font-size:11px"></div>
      <div style="font-size:10px;opacity:.6">BD Sensitive</div>
    </div>
  </div>
  <div id="spbanner" class="spbanner"><span id="spmsg"></span></div>
  <div class="nav">
    <button class="ntab on" data-n="0">Matrix</button>
    <button class="ntab" data-n="1">Vehicle Detail</button>
    <button class="ntab" data-n="2">Company Detail</button>
    <button class="ntab" data-n="3">Pursuit Decisions</button>
    <button class="ntab" data-n="4">Timeline</button>
    <button class="ntab" data-n="5">Intel Feed</button>
    <button class="ntab" data-n="6">Saved</button>
  </div>
  <div class="body">
    <div class="sidebar">
      <div class="sbs">
        <div class="sbt">Contract Vehicles</div>
        <div class="uzone" id="dz">
          <input type="file" id="fi" accept=".pdf,.txt,.doc,.docx" multiple style="display:none">
          <div style="font-size:18px;margin-bottom:3px">+</div>
          <div style="font-size:11px;color:#605e5c">Drop RFPs or <b id="fbtn" style="color:#1b3a6b;cursor:pointer">browse</b></div>
        </div>
        <button class="pastebtn" id="rpbtn">Paste RFP text</button>
        <div id="vlist" style="margin-top:5px"></div>
      </div>
      <div style="flex:1;display:flex;flex-direction:column;overflow:hidden;border-top:1px solid #edebe9">
        <div style="padding:8px 10px 4px;flex-shrink:0"><div class="sbt" style="margin-bottom:0">Companies</div></div>
        <div id="clist" style="flex:1;overflow-y:auto;padding:6px 8px"></div>
        <div style="padding:8px;border-top:1px solid #edebe9;flex-shrink:0">
          <button class="addbtn" id="acbtn">+ Add Company</button>
        </div>
      </div>
    </div>
    <div class="content" id="main">
      <div class="empty"><div style="font-size:44px;margin-bottom:10px">BD</div><h2>TWCG BD Intelligence Platform</h2><p>Loading your SharePoint data...</p></div>
    </div>
  </div>
</div>

<!-- Modals -->
<div class="overlay" id="rpModal">
  <div class="modal">
    <h2>Paste RFP Text</h2>
    <div class="hint">Paste RFP text below. Deuce will parse it automatically.</div>
    <div class="fr"><label>Vehicle Name</label><input id="rn" placeholder="e.g. MAPS"></div>
    <div class="fr"><label>Paste RFP text</label><textarea id="rt" style="height:180px" placeholder="Paste full RFP text here..."></textarea></div>
    <div class="pstat" id="rpstat"></div>
    <div class="mbtns"><button class="btn bc" id="rpcancel">Cancel</button><button class="btn bp" id="rpsubmit">Parse with Deuce</button></div>
  </div>
</div>

<div class="overlay" id="coModal">
  <div class="modal">
    <h2 id="coTitle">Add Company</h2>
    <div class="itbar">
      <button class="itab on" data-it="0">Manual</button>
      <button class="itab" data-it="1">Paste Text</button>
      <button class="itab" data-it="2">Upload File</button>
      <button class="itab" data-it="3">Website Text</button>
    </div>
    <div class="ipane on" id="ip0">
      <div class="f2"><div class="fr"><label>Company Name</label><input id="cn" placeholder="Full legal name"></div><div class="fr"><label>Short Name</label><input id="cs" placeholder="e.g. HII"></div></div>
      <div class="f2"><div class="fr"><label>CAGE Code</label><input id="cc" placeholder="5-char"></div><div class="fr"><label>UEI</label><input id="cu" placeholder="SAM.gov UEI"></div></div>
      <div class="f2"><div class="fr"><label>Icon (emoji)</label><input id="cico" value="*" maxlength="2"></div><div class="fr"><label>Headquarters</label><input id="chq" placeholder="City, State"></div></div>
      <div class="fr"><label>Business Size</label><select id="csz"><option>Small Business</option><option>Emerging Large Business</option><option>Large Business</option><option>Commercial-Sector Vendor</option></select></div>
      <div class="fr"><label>Business Types</label><div class="ckg" id="btg"></div></div>
      <div class="f2"><div class="fr"><label>NAICS Codes</label><input id="cna" placeholder="541611, 541512..."></div><div class="fr"><label>PSC Codes</label><input id="cps" placeholder="R408, D302..."></div></div>
      <div class="fr"><label>Facility Clearance</label><select id="ccl"><option>None</option><option>Confidential</option><option>Secret</option><option>Top Secret</option><option>TS/SCI</option></select></div>
      <div class="fr"><label>Certifications</label><div class="ckg" id="ceg"></div></div>
      <div class="fr"><label>CMMC Level</label><select id="cmmc"><option value="0">Not Started</option><option value="1">Level 1</option><option value="2">Level 2 Self</option><option value="2c">Level 2 C3PAO</option><option value="3">Level 3 DIBCAC</option></select></div>
      <div class="fr"><label>DCMA Systems</label><div class="ckg" id="dcg"></div></div>
      <div class="fr"><label>Primary Capabilities</label><input id="cap" placeholder="Program Mgmt, Cybersecurity..."></div>
      <div class="fr"><label>Website</label><div style="display:flex;gap:6px"><input id="cweb" placeholder="https://www.company.com"><button class="btn bp" id="scrapebtn" type="button" style="white-space:nowrap;padding:6px 12px;font-size:11px">Scrape with Deuce</button></div></div><div class="pstat" id="scrapestat"></div><div class="fr"><label>Annual Revenue</label><input id="rev" placeholder="e.g. $45M"></div>
      <div class="fr"><label>BD Notes</label><textarea id="nts" style="height:56px" placeholder="Competitive intel, relationship notes..."></textarea></div>
      <input type="hidden" id="eid">
      <div id="dnote" style="font-size:11px;color:#1b3a6b;background:#e8edf5;padding:7px 10px;border-radius:5px;margin-bottom:8px;display:none"></div>
      <div class="mbtns"><button class="btn bc" id="cocancel">Cancel</button><button class="btn bp" id="cosave">Save Company</button></div>
    </div>
    <div class="ipane" id="ip1">
      <div class="hint">Paste SAM.gov page text, GovWin export, or capability statement. Deuce extracts the intel.</div>
      <div class="fr"><label>Company Name (helps Deuce focus)</label><input id="pcn" placeholder="e.g. HII Mission Technologies"></div>
      <div class="fr"><label>Paste source text</label><textarea id="ptxt" style="height:200px" placeholder="Paste SAM.gov, GovWin, capability statement..."></textarea></div>
      <div class="pstat" id="pstat"></div>
      <div class="mbtns"><button class="btn bc" id="pcancel">Cancel</button><button class="btn bp" id="pparse">Parse with Deuce</button></div>
    </div>
    <div class="ipane" id="ip2">
      <div class="hint">Upload a capability statement PDF or GovWin export.</div>
      <div class="fr"><label>Company Name (optional)</label><input id="ucn" placeholder="e.g. Booz Allen Hamilton"></div>
      <div class="uzone" id="cdz" style="padding:20px">
        <input type="file" id="cfi" accept=".pdf,.txt,.doc,.docx" style="display:none">
        <div style="font-size:24px;margin-bottom:6px">+</div>
        <div style="font-size:12px;color:#605e5c">Drop file or <b id="cfbtn" style="color:#1b3a6b;cursor:pointer">browse</b></div>
      </div>
      <div id="ufn" style="font-size:11px;color:#107c10;margin-top:5px;display:none"></div>
      <div class="pstat" id="ustat"></div>
      <div class="mbtns"><button class="btn bc" id="ucancel">Cancel</button><button class="btn bp" id="uparse">Parse with Deuce</button></div>
    </div>
  </div>
</div>

<div class="ipane" id="ip3">
      <div class="hint">Copy all text from a SAM.gov profile, GovWin page, or any website (Ctrl+A, Ctrl+C) and paste below. Deuce will extract the intel.</div>
      <div class="fr"><label>Company Name (optional)</label><input id="wcn" placeholder="e.g. Booz Allen Hamilton"></div>
      <div class="fr"><label>Paste website text</label><textarea id="wtxt" style="height:200px" placeholder="Paste copied website text here..."></textarea></div>
      <div class="pstat" id="wstat"></div>
      <div class="mbtns"><button class="btn bc" id="wcancel">Cancel</button><button class="btn bp" id="wparse">Parse with Deuce</button></div>
    </div>
  </div>
</div>

<div class="overlay" id="enrichModal">
  <div class="modal" style="max-width:700px">
    <h2>Enrich Company Profile</h2>
    <div class="hint">Deuce found new intel. Review each field and choose which value to keep. Empty fields are auto-filled.</div>
    <div id="enrichDiff" style="margin-bottom:14px"></div>
    <div class="mbtns"><button class="btn bc" id="enrichcancel">Cancel</button><button class="btn bp" id="enrichsave">Apply Selected</button></div>
  </div>
</div>

<div class="overlay" id="enrichSrcModal">
  <div class="modal" style="max-width:580px">
    <h2>Enrich with New Source</h2>
    <div class="itbar">
      <button class="itab on" data-et="0">Paste Text</button>
      <button class="itab" data-et="1">Upload File</button>
      <button class="itab" data-et="2">Website Text</button>
    </div>
    <div class="epane on" id="ep0">
      <div class="hint">Paste SAM.gov, GovWin, or capability statement text. Deuce merges with existing profile.</div>
      <div class="fr"><label>Paste source text</label><textarea id="eptxt" style="height:200px" placeholder="Paste text here..."></textarea></div>
      <div class="pstat" id="epstat"></div>
      <div class="mbtns"><button class="btn bc" id="epcancel">Cancel</button><button class="btn bp" id="epparse">Parse with Deuce</button></div>
    </div>
    <div class="epane" id="ep1">
      <div class="hint">Upload a capability statement PDF or GovWin export.</div>
      <div class="uzone" id="ecdz" style="padding:20px">
        <input type="file" id="ecfi" accept=".pdf,.txt,.doc,.docx" style="display:none">
        <div style="font-size:24px;margin-bottom:6px">+</div>
        <div style="font-size:12px;color:#605e5c">Drop file or <b id="ecfbtn" style="color:#1b3a6b;cursor:pointer">browse</b></div>
      </div>
      <div id="ecfn" style="font-size:11px;color:#107c10;margin-top:5px;display:none"></div>
      <div class="pstat" id="ecstat"></div>
      <div class="mbtns"><button class="btn bc" id="ecucancel">Cancel</button><button class="btn bp" id="ecuparse">Parse with Deuce</button></div>
    </div>
    <div class="epane" id="ep2">
      <div class="hint">Copy all text from any website (Ctrl+A, Ctrl+C) and paste below.</div>
      <div class="fr"><label>Paste website text</label><textarea id="ewtxt" style="height:200px" placeholder="Paste copied website text here..."></textarea></div>
      <div class="pstat" id="ewstat"></div>
      <div class="mbtns"><button class="btn bc" id="ewcancel">Cancel</button><button class="btn bp" id="ewparse">Parse with Deuce</button></div>
    </div>
  </div>
</div>

<div class="overlay" id="cellModal"><div class="modal" style="max-width:560px"><div id="cellContent"></div></div></div>
<div class="overlay" id="svModal">
  <div class="modal" style="max-width:380px">
    <h2>Save Comparison</h2>
    <div class="fr"><label>Name</label><input id="svn" placeholder="e.g. MAPS Q2 2026"></div>
    <div class="fr"><label>Notes</label><textarea id="svnt" style="height:56px"></textarea></div>
    <div class="mbtns"><button class="btn bc" id="svcancel">Cancel</button><button class="btn bp" id="svsave">Save</button></div>
  </div>
</div>
<div class="overlay" id="dtModal">
  <div class="modal" style="max-width:320px">
    <h2 id="dtTitle">Set Date</h2>
    <div class="fr"><label>Date</label><input type="date" id="dtInput"></div>
    <div class="mbtns"><button class="btn bc" id="dtcancel">Cancel</button><button class="btn bp" id="dtsave">Set Date</button></div>
  </div>
</div>`;
  }

  // ── BOOT ──────────────────────────────────────────────────────────────────────
  private _boot(): void {
    const u = this.domElement.querySelector('#upill') as HTMLElement;
    if (u) u.textContent = this.context.pageContext.user.displayName;
    this._setupNav();
    this._setupSidebar();
    this._setupModals();
    this._renderSidebar();
    this._renderContent();
    this._loadSP();
  }

  private _q(sel: string): HTMLElement|null { return this.domElement.querySelector(sel); }
  private _qa(sel: string): NodeListOf<Element> { return this.domElement.querySelectorAll(sel); }

  private _setNav(n: number) {
    this.nav = n;
    this._qa('.ntab').forEach((t, i) => t.classList.toggle('on', i === n));
  }

  private _setupNav(): void {
    this._qa('.ntab').forEach(tab => {
      tab.addEventListener('click', () => {
        this._setNav(parseInt(tab.getAttribute('data-n') || '0'));
        this._renderContent();
      });
    });
  }

  private _setupSidebar(): void {
    this._q('#fbtn')?.addEventListener('click', () => (this._q('#fi') as HTMLInputElement)?.click());
    this._q('#rpbtn')?.addEventListener('click', () => this._q('#rpModal')?.classList.add('on'));
    this._q('#acbtn')?.addEventListener('click', () => this._openCoModal(null));
    this._q('#cfbtn')?.addEventListener('click', () => (this._q('#cfi') as HTMLInputElement)?.click());
    const dz = this._q('#dz');
    const fi = this._q('#fi') as HTMLInputElement;
    dz?.addEventListener('dragover', e => e.preventDefault());
    dz?.addEventListener('drop', e => { e.preventDefault(); this._handleRfpFiles((e as DragEvent).dataTransfer?.files); });
    fi?.addEventListener('change', () => { this._handleRfpFiles(fi.files); fi.value = ''; });
    const cdz = this._q('#cdz');
    const cfi = this._q('#cfi') as HTMLInputElement;
    cdz?.addEventListener('dragover', e => e.preventDefault());
    cdz?.addEventListener('drop', e => { e.preventDefault(); this._handleCoFile((e as DragEvent).dataTransfer?.files?.[0]); });
    cfi?.addEventListener('change', () => { this._handleCoFile(cfi.files?.[0]); });
  }

  private _setupModals(): void {
    this._q('#rpcancel')?.addEventListener('click', () => this._q('#rpModal')?.classList.remove('on'));
    this._q('#rpsubmit')?.addEventListener('click', () => this._submitRfp());
    this._q('#cocancel')?.addEventListener('click', () => this._closeCoModal());
    this._q('#cosave')?.addEventListener('click', () => this._saveCompany());
    this._q('#pcancel')?.addEventListener('click', () => this._closeCoModal());
    this._q('#pparse')?.addEventListener('click', () => this._parseFromPaste());
    this._q('#ucancel')?.addEventListener('click', () => this._closeCoModal());
    this._q('#uparse')?.addEventListener('click', () => this._parseFromUpload());
    this._q('#svcancel')?.addEventListener('click', () => this._q('#svModal')?.classList.remove('on'));
    this._q('#svsave')?.addEventListener('click', () => this._saveComp());
    this._q('#dtcancel')?.addEventListener('click', () => this._q('#dtModal')?.classList.remove('on'));
    this._q('#wcancel')?.addEventListener('click', () => this._closeCoModal());
    this._q('#scrapebtn')?.addEventListener('click', () => this._scrapeWebsite());
    this._q('#wparse')?.addEventListener('click', () => this._parseFromWebsite());
    this._q('#enrichcancel')?.addEventListener('click', () => this._q('#enrichModal')?.classList.remove('on'));
    this._q('#enrichsave')?.addEventListener('click', () => this._applyEnrich());
    this._q('#epcancel')?.addEventListener('click', () => this._q('#enrichSrcModal')?.classList.remove('on'));
    this._q('#epparse')?.addEventListener('click', () => this._enrichFromPaste());
    this._q('#ecucancel')?.addEventListener('click', () => this._q('#enrichSrcModal')?.classList.remove('on'));
    this._q('#ecuparse')?.addEventListener('click', () => this._enrichFromUpload());
    this._q('#ewcancel')?.addEventListener('click', () => this._q('#enrichSrcModal')?.classList.remove('on'));
    this._q('#ewparse')?.addEventListener('click', () => this._enrichFromWebsite());
    this._q('#ecfbtn')?.addEventListener('click', () => (this._q('#ecfi') as HTMLInputElement)?.click());
    const ecdz = this._q('#ecdz'); const ecfi = this._q('#ecfi') as HTMLInputElement;
    ecdz?.addEventListener('dragover', e => e.preventDefault());
    ecdz?.addEventListener('drop', e => { e.preventDefault(); this.enrichFile = (e as DragEvent).dataTransfer?.files?.[0] || null; if(this.enrichFile){const fn=this._q('#ecfn') as HTMLElement;fn.style.display='block';fn.textContent=this.enrichFile.name;} });
    ecfi?.addEventListener('change', () => { this.enrichFile = ecfi.files?.[0] || null; if(this.enrichFile){const fn=this._q('#ecfn') as HTMLElement;fn.style.display='block';fn.textContent=this.enrichFile.name;} });
    this._qa('[data-et]').forEach(tab => {
      tab.addEventListener('click', () => {
        const n = parseInt(tab.getAttribute('data-et') || '0');
        this._qa('[data-et]').forEach((t,i) => { t.classList.toggle('on',i===n); (t as HTMLElement).style.background=i===n?'#1b3a6b':'transparent'; (t as HTMLElement).style.color=i===n?'#fff':'#605e5c'; });
        this._qa('.epane').forEach((p,i) => p.classList.toggle('on',i===n));
      });
    });
    this._q('#dtsave')?.addEventListener('click', () => this._saveDt());
    this._qa('.itab').forEach(tab => {
      tab.addEventListener('click', () => {
        const n = parseInt(tab.getAttribute('data-it') || '0');
        this._qa('.itab').forEach((t, i) => { t.classList.toggle('on', i===n); (t as HTMLElement).style.background = i===n?'#1b3a6b':'transparent'; (t as HTMLElement).style.color = i===n?'#fff':'#605e5c'; });
        this._qa('.ipane').forEach((p, i) => p.classList.toggle('on', i===n));
      });
    });
  }

  private _renderSidebar(): void {
    const sm: any = {parsed:'sok',parsing:'swt',raw:'srw'};
    const sl: any = {parsed:'Ready',parsing:'Parsing...',raw:'Review'};
    const vl = this._q('#vlist');
    if (vl) {
      const vk = Object.keys(this.V);
      vl.innerHTML = vk.length ? vk.map(id => {
        const v = this.V[id];
        return `<div class="sitem${id===this.aV?' on':''}" data-vid="${id}"><div class="sico" style="background:#e8edf5">RFP</div><div class="sinf"><div class="snm">${v.dn}</div><div class="smt">${Math.round(v.size/1024)}KB</div><span class="stag ${sm[v.status]}">${sl[v.status]}</span></div><span class="dx" data-dvid="${id}">X</span></div>`;
      }).join('') : '<div style="font-size:10px;color:#a19f9d;padding:4px">No vehicles loaded</div>';
    }
    const cl = this._q('#clist');
    if (cl) {
      const ck = Object.keys(this.C);
      cl.innerHTML = ck.length ? ck.map(id => {
        const c = this.C[id];
        return `<div class="sitem${id===this.aC?' on':''}" data-cid="${id}"><div class="sico" style="font-size:16px;background:#f3f2f1">${c.icon}</div><div class="sinf"><div class="snm">${c.name}</div><div class="smt">${c.size}</div></div><span class="dx" data-dcid="${id}">X</span></div>`;
      }).join('') : '<div style="font-size:10px;color:#a19f9d;padding:4px">Click Add Company below</div>';
    }
    this._qa('[data-vid]').forEach(el => el.addEventListener('click', () => { this.aV = el.getAttribute('data-vid'); this._setNav(1); this._renderSidebar(); this._renderContent(); }));
    this._qa('[data-cid]').forEach(el => el.addEventListener('click', () => { this.aC = el.getAttribute('data-cid'); this._setNav(2); this._renderSidebar(); this._renderContent(); }));
    this._qa('[data-dvid]').forEach(el => el.addEventListener('click', e => { e.stopPropagation(); const id = el.getAttribute('data-dvid')!; delete this.V[id]; delete this.M[id]; if (this.aV===id) this.aV=null; this._renderSidebar(); this._renderContent(); }));
    this._qa('[data-dcid]').forEach(el => el.addEventListener('click', e => { e.stopPropagation(); const id = el.getAttribute('data-dcid')!; delete this.C[id]; Object.keys(this.M).forEach(vid => delete this.M[vid][id]); if (this.aC===id) this.aC=null; this._renderSidebar(); this._renderContent(); }));
  }

  private _renderContent(): void {
    this._ensureMx();
    if (this.nav===0) this._renderMatrix();
    else if (this.nav===1) this._renderVD();
    else if (this.nav===2) this._renderCD();
    else if (this.nav===3) this._renderPD();
    else if (this.nav===4) this._renderTL();
    else if (this.nav===5) this._renderIF();
    else if (this.nav===6) this._renderSV();
  }

  private _renderMatrix(): void {
    const mc = this._q('#main')!;
    const vids = Object.keys(this.V), cids = Object.keys(this.C);
    if (!vids.length && !cids.length) { mc.innerHTML = '<div class="empty"><div style="font-size:40px;margin-bottom:8px">BD</div><h2>Build Your Matrix</h2><p>Upload RFPs and add companies to generate your pursuit comparison.</p></div>'; return; }
    const th = vids.map(vid => `<th data-hv="${vid}" style="cursor:pointer">${this.V[vid].dn}</th>`).join('');
    const rows = cids.map(cid => {
      const co = this.C[cid];
      const cells = vids.map(vid => {
        const v = this.V[vid];
        if (v.status==='parsing') return '<td><div style="color:#a19f9d;font-size:10px">Parsing...</div></td>';
        const pw = this.M[vid][cid].po != null ? this.M[vid][cid].po : this._cpwin(vid,cid);
        const vs = this.M[vid][cid].so != null ? this.M[vid][cid].so : this._cscore(vid,cid);
        const pu = this.M[vid][cid].pursue, st = this.M[vid][cid].stage||'Tracking', c = this._sc(pw);
        return `<td data-cv="${vid}" data-cc="${cid}" style="cursor:pointer"><div class="ci"><div class="cpw" style="color:${c}">${pw}%</div>${vs!=null?`<div style="font-size:10px;color:#605e5c">${vs.toLocaleString()} pts</div>`:''}<div class="cbar"><div class="cbarf" style="width:${pw}%;background:${c}"></div></div><span class="pill ${pu==='pursue'?'go':pu==='conditional'?'vf':'ac'}">${this._dl(pu)}</span><span class="spill" style="background:${this._sbg(st)};color:${this._sfg(st)}">${st}</span></div></td>`;
      }).join('');
      return `<tr><td data-hc="${cid}" style="cursor:pointer"><div style="display:flex;align-items:center;gap:5px"><span style="font-size:15px">${co.icon}</span><div><div style="font-weight:600;font-size:12px">${co.name}</div><div style="font-size:10px;color:#605e5c">${co.size}</div></div></div></td>${cells}</tr>`;
    }).join('');
    mc.innerHTML = `<div class="hrow"><div class="htitle">Pursuit Matrix <span style="font-size:12px;font-weight:400;color:#605e5c">${cids.length} companies - ${vids.length} vehicles</span></div><div style="display:flex;gap:7px"><button class="btn bg" id="rabtn">Reset All</button><button class="btn bp" id="svobtn">Save Comparison</button></div></div><div class="mxw"><table class="mx"><thead><tr><th>Company / Vehicle</th>${th}</tr></thead><tbody>${rows||'<tr><td colspan="99" style="text-align:center;color:#a19f9d;padding:16px">Add companies using the button in the sidebar</td></tr>'}</tbody></table></div><div style="display:flex;gap:12px;font-size:10px;color:#605e5c;margin-top:6px;flex-wrap:wrap"><span style="color:#107c10;font-weight:600">Pursue 55%+</span><span style="color:#835c00;font-weight:600">Conditional 40-54%</span><span style="color:#a4262c;font-weight:600">No-Go under 40%</span><span>Click cell to edit</span></div>`;
    this._qa('[data-cv]').forEach(el => el.addEventListener('click', () => this._openCell(el.getAttribute('data-cv')!, el.getAttribute('data-cc')!)));
    this._qa('[data-hv]').forEach(el => el.addEventListener('click', () => { this.aV = el.getAttribute('data-hv'); this._setNav(1); this._renderSidebar(); this._renderContent(); }));
    this._qa('[data-hc]').forEach(el => el.addEventListener('click', () => { this.aC = el.getAttribute('data-hc'); this._setNav(2); this._renderSidebar(); this._renderContent(); }));
    this._q('#rabtn')?.addEventListener('click', () => { if (!confirm('Reset all overrides?')) return; Object.keys(this.M).forEach(vid => Object.keys(this.M[vid]).forEach(cid => { this.M[vid][cid].po=null; this.M[vid][cid].so=null; const fs=this.V[vid].parsed?.pwinFactors||[]; fs.forEach((f:any)=>{ if(this.PW[vid]?.[cid]) this.PW[vid][cid][f.key]=f.val; }); })); this._renderContent(); });
    this._q('#svobtn')?.addEventListener('click', () => this._q('#svModal')?.classList.add('on'));
  }

  private _openCell(vid: string, cid: string): void {
    const v = this.V[vid], co = this.C[cid], p = v.parsed;
    const s = this.PW[vid]?.[cid] || {};
    const pw = this._cpwin(vid,cid), vs = this._cscore(vid,cid), cell = this.M[vid][cid];
    const dispPw = cell.po != null ? cell.po : pw;
    const stages = ['Tracking','Shaping','Bid','No-Bid','Award','Loss'];
    const fs = p?.pwinFactors || [];
    const cc = this._q('#cellContent')!;
    cc.innerHTML = `<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px"><div><div style="font-size:14px;font-weight:700;color:#1b3a6b">${co.name} x ${v.dn}</div><div style="font-size:11px;color:#605e5c;margin-top:2px">${co.size}</div></div><div style="font-size:22px;font-weight:700;color:${this._sc(pw)}" id="cpd">${dispPw}%</div></div><div style="margin-bottom:10px"><div class="stitle">Pipeline Stage</div><div style="display:flex;gap:5px;flex-wrap:wrap" id="stg">${stages.map(st=>`<button class="btn ${cell.stage===st?'bp':'bg'}" style="padding:4px 10px;font-size:11px" data-st="${st}">${st}</button>`).join('')}</div></div><div style="margin-bottom:10px"><div class="stitle">Pursuit Decision</div><div style="display:flex;gap:5px" id="dec">${['pursue','conditional','no-go'].map(d=>`<button class="btn ${cell.pursue===d?'bp':'bg'}" style="padding:4px 10px;font-size:11px" data-dc="${d}">${this._dl(d)}</button>`).join('')}</div></div><div style="background:#f3f2f1;border-radius:7px;padding:10px;margin-bottom:10px"><div class="stitle">P-Win Factors</div>${fs.map((f:any)=>`<div class="slrow" style="margin-bottom:6px"><div class="sli"><div class="sll">${f.label}</div><div class="sld">${f.desc}</div></div><div class="opts" data-fk="${f.key}" data-fc="${f.color}">${[1,2,3].map(v2=>{const on=(s[f.key]||1)===v2;return`<button class="opt${on?' on':''}" style="${on?`background:${f.color};border-color:${f.color};color:#fff`:''}" data-fv="${v2}">${['Low','Mod','High'][v2-1]}</button>`;}).join('')}</div></div>`).join('')}<div class="ovr"><label>P-Win Override:</label><input type="number" min="0" max="100" id="pov" value="${cell.po!=null?cell.po:''}" placeholder="--"><span style="font-size:11px;color:#605e5c">% blank=calculated</span><button class="btn bg" style="padding:3px 8px;font-size:11px" id="apwbtn">Apply</button></div></div>${vs!=null?`<div class="ovr" style="margin-bottom:10px;background:#fff;border:1px solid #edebe9"><label>Score Override:</label><input type="number" min="0" id="sov" value="${cell.so!=null?cell.so:''}" placeholder="--"><span style="font-size:11px;color:#605e5c">pts blank=calc: ${vs.toLocaleString()}</span><button class="btn bg" style="padding:3px 8px;font-size:11px" id="asbtn">Apply</button></div>`:''}<div style="margin-bottom:10px"><div class="stitle">Intel Note</div><textarea id="cnote" style="width:100%;padding:6px 8px;border:1px solid #edebe9;border-radius:5px;font-size:12px;height:52px;resize:none;font-family:Segoe UI,sans-serif">${cell.notes||''}</textarea></div><div style="display:flex;gap:7px;justify-content:space-between"><button class="btn bd" id="crst">Reset Cell</button><div style="display:flex;gap:7px"><button class="btn bc" id="ccancel">Cancel</button><button class="btn bp" id="csavebtn">Save</button></div></div>`;
    this._q('#cellModal')?.classList.add('on');
    this._qa('#stg [data-st]').forEach(btn => btn.addEventListener('click', () => { cell.stage = btn.getAttribute('data-st')!; this._qa('#stg [data-st]').forEach(b => { b.className='btn bg'; (b as HTMLElement).style.cssText='padding:4px 10px;font-size:11px'; }); btn.className='btn bp'; (btn as HTMLElement).style.cssText='padding:4px 10px;font-size:11px'; }));
    this._qa('#dec [data-dc]').forEach(btn => btn.addEventListener('click', () => { cell.pursue = btn.getAttribute('data-dc')!; this._qa('#dec [data-dc]').forEach(b => { b.className='btn bg'; (b as HTMLElement).style.cssText='padding:4px 10px;font-size:11px'; }); btn.className='btn bp'; (btn as HTMLElement).style.cssText='padding:4px 10px;font-size:11px'; }));
    this._qa('.opts[data-fk]').forEach(grp => {
      const key = grp.getAttribute('data-fk')!, color = grp.getAttribute('data-fc')!;
      grp.querySelectorAll('[data-fv]').forEach(btn => btn.addEventListener('click', () => {
        const val = parseInt(btn.getAttribute('data-fv')!);
        if (!this.PW[vid]) this.PW[vid] = {}; if (!this.PW[vid][cid]) this.PW[vid][cid] = {};
        this.PW[vid][cid][key] = val;
        grp.querySelectorAll('[data-fv]').forEach(b => { b.className='opt'; (b as HTMLElement).style.cssText=''; });
        btn.className='opt on'; (btn as HTMLElement).style.background=color; (btn as HTMLElement).style.borderColor=color; (btn as HTMLElement).style.color='#fff';
        const np = this._cpwin(vid,cid), d = this._q('#cpd'); if (d) { d.textContent=np+'%'; d.style.color=this._sc(np); }
      }));
    });
    this._q('#apwbtn')?.addEventListener('click', () => { const v2 = (this._q('#pov') as HTMLInputElement).value; cell.po = v2===''?null:parseInt(v2); const np=cell.po!=null?cell.po:this._cpwin(vid,cid); const d=this._q('#cpd'); if(d){d.textContent=np+'%';d.style.color=this._sc(np);} });
    this._q('#asbtn')?.addEventListener('click', () => { const v2=(this._q('#sov') as HTMLInputElement).value; cell.so=v2===''?null:parseInt(v2); });
    this._q('#crst')?.addEventListener('click', () => { this.M[vid][cid]={pursue:'conditional',notes:'',stage:'Tracking',po:null,so:null}; const fs2=this.V[vid].parsed?.pwinFactors||[]; fs2.forEach((f:any)=>{if(this.PW[vid]?.[cid])this.PW[vid][cid][f.key]=f.val;}); this._q('#cellModal')?.classList.remove('on'); this._renderContent(); });
    this._q('#ccancel')?.addEventListener('click', () => this._q('#cellModal')?.classList.remove('on'));
    this._q('#csavebtn')?.addEventListener('click', () => { const n=(this._q('#cnote') as HTMLTextAreaElement)?.value.trim(); if(n&&n!==cell.notes){cell.notes=n;this._addLog('cell',vid+'_'+cid,n);} this._saveMx(); this._q('#cellModal')?.classList.remove('on'); this._renderContent(); });
  }

  private _renderVD(): void {
    const mc = this._q('#main')!;
    if (!this.aV||!this.V[this.aV]) { mc.innerHTML='<div class="empty"><h2>Select a Vehicle</h2><p>Click a vehicle in the sidebar.</p></div>'; return; }
    const v = this.V[this.aV];
    if (v.status==='parsing') { mc.innerHTML=`<div style="text-align:center;padding:40px;color:#1b3a6b;font-weight:600">Deuce is analyzing ${v.dn}...</div>`; return; }
    const p = v.parsed;
    const dfs = [{key:'releaseDate',lbl:'Release Date'},{key:'dueDate',lbl:'Proposals Due'},{key:'awardDate',lbl:'Expected Award'}];
    const dateCards = dfs.map(df => { const val=(v.de?.[df.key])||p[df.key]||'TBD'; const tbd=val==='TBD'; return `<div class="sc2"><div class="sl">${df.lbl}</div><div class="sv" style="color:${tbd?'#a4262c':'#323130'}">${val}</div>${tbd?`<button class="btn bg" style="padding:2px 8px;font-size:10px;margin-top:4px" data-pd="${df.key}" data-pl="${df.lbl}">Set estimate</button>`:''}</div>`; }).join('');
    const tabLabels = ['Gates','Domains','Score Ref','Company Compare'];
    mc.innerHTML = `<div class="hrow"><div><div class="htitle">${p.vehicleName}</div><div style="font-size:11px;color:#605e5c;margin-top:2px">${p.agency} - ${p.type} - ${p.pop}</div><div style="display:flex;gap:5px;flex-wrap:wrap;margin-top:5px">${p.solNum&&p.solNum!=='TBD'?`<span class="pill sc">Sol# ${p.solNum}</span>`:''} ${p.ceiling&&p.ceiling!=='TBD'?`<span class="pill go">Ceiling: ${p.ceiling}</span>`:''}</div></div></div><div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:12px">${dateCards}</div><div class="dtabs">${tabLabels.map((t,i)=>`<button class="dtab${i===0?' on':''}" data-vt="${i}">${t}</button>`).join('')}</div><div class="dpane on" id="vp0">${this._bGates(p)}</div><div class="dpane" id="vp1">${this._bDoms(p)}</div><div class="dpane" id="vp2">${this._bSRef(p)}</div><div class="dpane" id="vp3">${this._bVCC(this.aV!)}</div>`;
    this._qa('[data-vt]').forEach(el => el.addEventListener('click', () => { const i=parseInt(el.getAttribute('data-vt')!); this._qa('[data-vt]').forEach((t,j)=>t.classList.toggle('on',j===i)); this._qa('[id^="vp"]').forEach((pane,j)=>pane.classList.toggle('on',j===i)); }));
    this._qa('[data-pd]').forEach(el => el.addEventListener('click', () => { this.dtCb={vid:this.aV,field:el.getAttribute('data-pd')}; (this._q('#dtTitle') as HTMLElement).textContent='Set Date - '+el.getAttribute('data-pl'); (this._q('#dtInput') as HTMLInputElement).value=''; this._q('#dtModal')?.classList.add('on'); }));
    this._qa('[data-cv]').forEach(el => el.addEventListener('click', () => this._openCell(el.getAttribute('data-cv')!, el.getAttribute('data-cc')!)));
  }

  private _bGates(p: any): string { return (p.gates||[]).map((g:any)=>`<div class="card"><div style="display:flex;align-items:center;gap:9px;padding:11px 14px"><div style="width:22px;height:22px;border-radius:50%;background:${g.color};display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#fff;flex-shrink:0">${g.num}</div><div style="font-size:13px;font-weight:600;flex:1">${g.label}</div></div><div style="padding:0 14px 12px"><table class="t"><thead><tr><th style="width:28%">Criterion</th><th>Requirement</th><th style="width:70px">Status</th></tr></thead><tbody>${(g.rows||[]).map((r:any)=>`<tr><td style="font-weight:500">${r[0]}</td><td style="color:#605e5c;font-size:11px">${r[1]}</td><td><span class="pill ${r[3]||'as'}">${r[2]}</span></td></tr>`).join('')}</tbody></table></div></div>`).join(''); }
  private _bDoms(p: any): string { return `<div class="dg">${(p.domains||[]).map((d:any)=>`<div class="dc"><div class="dt"><div class="dn">${d.name}</div><div style="font-size:10px;color:#605e5c;margin-top:1px">NAICS ${d.naics} - ${d.note}</div><div class="bb"><div class="bf" style="width:${d.pct}%;background:${d.color}"></div></div><div style="font-size:10px;color:${d.color};font-weight:600">${d.pct}% coverage</div><div style="font-size:10px;color:${d.rec?'#107c10':'#a4262c'};font-weight:600;margin-top:2px">${d.rec?'Recommended':'Not recommended'}</div></div>${(d.caps||[]).map((cap:any)=>`<div class="cr"><span>${cap[0]}</span><span class="${cap[1]==='aligned'?'ali':cap[1]==='partial'?'par':'gap'}">${cap[1].charAt(0).toUpperCase()+cap[1].slice(1)}</span></div>`).join('')}</div>`).join('')}</div>`; }
  private _bSRef(p: any): string { if (!p.scoreRef?.length) return '<div style="padding:14px;color:#605e5c;font-size:12px">No scoring data.</div>'; return `<table class="t"><thead><tr><th>Category</th><th style="width:100px">Max Points</th><th>Key Levers</th></tr></thead><tbody>${p.scoreRef.map((row:any,i:number)=>`<tr style="${i===p.scoreRef.length-1?'font-weight:700;background:#faf9f8':''}"><td>${row[0]}</td><td>${row[1]}</td><td style="color:#605e5c;font-size:11px">${row[2]||''}</td></tr>`).join('')}</tbody></table>`; }
  private _bVCC(vid: string): string {
    const cids = Object.keys(this.C); if (!cids.length) return '<div style="padding:14px;color:#605e5c;font-size:12px">No companies added.</div>';
    this._ensureMx();
    return `<table class="t"><thead><tr><th>Company</th><th>P-Win</th><th>Score</th><th>Decision</th><th>Stage</th><th></th></tr></thead><tbody>${cids.map(cid=>{ const co=this.C[cid]; const pw=this.M[vid][cid].po!=null?this.M[vid][cid].po:this._cpwin(vid,cid); const vs=this.M[vid][cid].so!=null?this.M[vid][cid].so:this._cscore(vid,cid); const pu=this.M[vid][cid].pursue,st=this.M[vid][cid].stage||'Tracking',c=this._sc(pw); return `<tr><td><div style="display:flex;align-items:center;gap:5px"><span>${co.icon}</span><span style="font-weight:600">${co.name}</span></div></td><td><span style="font-size:14px;font-weight:700;color:${c}">${pw}%</span></td><td style="font-size:11px">${vs!=null?vs.toLocaleString()+' pts':'-'}</td><td><span class="pill ${this._dp(pu)}">${this._dl(pu)}</span></td><td><span style="font-size:10px;font-weight:700;padding:2px 6px;border-radius:8px;background:${this._sbg(st)};color:${this._sfg(st)}">${st}</span></td><td><button class="btn bg" style="padding:3px 8px;font-size:10px" data-cv="${vid}" data-cc="${cid}">Edit</button></td></tr>`; }).join('')}</tbody></table>`;
  }

  private _renderCD(): void {
    const mc = this._q('#main')!;
    if (!this.aC||!this.C[this.aC]) { mc.innerHTML='<div class="empty"><h2>Select a Company</h2><p>Click a company in the sidebar.</p></div>'; return; }
    const co = this.C[this.aC];
    const vids = Object.keys(this.V).filter(v => this.V[v].status!=='parsing');
    this._ensureMx();
    mc.innerHTML = `<div class="hrow"><div style="display:flex;align-items:center;gap:10px"><div style="font-size:28px">${co.icon}</div><div><div class="htitle">${co.name}</div><div style="font-size:11px;color:#605e5c">${co.size}${co.hq?' - '+co.hq:''}</div></div></div><button class="btn bg" id="editco">Edit Profile</button><button class="btn bp" id="enrichco" style="margin-left:6px">+ Enrich with Source</button></div><div class="sgrid"><div class="sc2"><div class="sl">Business Types</div><div style="display:flex;flex-wrap:wrap;margin-top:3px">${(co.bizTypes||[]).map((t:string)=>`<span class="tag" style="background:#e8edf5;color:#1b3a6b">${t}</span>`).join('')||'-'}</div></div><div class="sc2"><div class="sl">Clearance</div><div class="sv" style="color:${co.clearance==='None'?'#a4262c':'#107c10'}">${co.clearance||'None'}</div></div><div class="sc2"><div class="sl">CMMC</div><div class="sv">${({0:'Not Started',1:'Level 1',2:'Level 2 Self','2c':'Level 2 C3PAO',3:'Level 3 DIBCAC'} as any)[co.cmmc]||'-'}</div></div><div class="sc2"><div class="sl">NAICS</div><div style="font-size:11px">${(co.naics||[]).join(', ')||'-'}</div></div><div class="sc2"><div class="sl">PSC</div><div style="font-size:11px">${(co.psc||[]).join(', ')||'-'}</div></div><div class="sc2"><div class="sl">Certifications</div><div style="display:flex;flex-wrap:wrap;margin-top:3px">${(co.certs||[]).map((t:string)=>`<span class="tag" style="background:#dff6dd;color:#107c10">${t}</span>`).join('')||'-'}</div></div></div>${co.notes?`<div class="card" style="padding:12px;margin-bottom:10px"><div class="stitle">BD Notes</div><div style="font-size:12px;line-height:1.6">${co.notes}</div></div>`:''}<div class="stitle">Vehicle Pursuit Summary</div>${vids.length?`<table class="t"><thead><tr><th>Vehicle</th><th>P-Win</th><th>Score</th><th>Decision</th><th>Stage</th><th></th></tr></thead><tbody>${vids.map(vid=>{ const pw=this.M[vid][this.aC!].po!=null?this.M[vid][this.aC!].po:this._cpwin(vid,this.aC!); const vs=this.M[vid][this.aC!].so!=null?this.M[vid][this.aC!].so:this._cscore(vid,this.aC!); const pu=this.M[vid][this.aC!].pursue,st=this.M[vid][this.aC!].stage||'Tracking',c=this._sc(pw); return `<tr><td style="font-weight:600">${this.V[vid].dn}</td><td><span style="font-size:14px;font-weight:700;color:${c}">${pw}%</span></td><td style="font-size:11px">${vs!=null?vs.toLocaleString()+' pts':'-'}</td><td><span class="pill ${this._dp(pu)}">${this._dl(pu)}</span></td><td><span style="font-size:10px;font-weight:700;padding:2px 6px;border-radius:8px;background:${this._sbg(st)};color:${this._sfg(st)}">${st}</span></td><td><button class="btn bg" style="padding:3px 8px;font-size:10px" data-cv="${vid}" data-cc="${this.aC}">Edit</button></td></tr>`; }).join('')}</tbody></table>`:'<div style="font-size:12px;color:#605e5c;padding:12px">No vehicles uploaded yet.</div>'}<div style="margin-top:14px"><div class="stitle">Intel Log</div>${this.logs.filter(l=>l.scope==='company'&&l.sid===this.aC).slice(-8).reverse().map(l=>`<div class="log"><div class="logm">${l.author} - ${l.date}</div>${l.entry}</div>`).join('')||'<div style="font-size:11px;color:#a19f9d">No entries yet.</div>'}<div style="display:flex;gap:6px;margin-top:8px"><textarea id="cli" style="flex:1;padding:6px 8px;border:1px solid #edebe9;border-radius:5px;font-size:12px;height:52px;resize:none;font-family:Segoe UI,sans-serif" placeholder="Add intel note..."></textarea><button class="btn bp" id="addcl">Add</button></div></div>`;
    this._q('#editco')?.addEventListener('click', () => this._openCoModal(this.aC));
    this._q('#enrichco')?.addEventListener('click', () => this._openEnrichModal(this.aC!));
    this._q('#addcl')?.addEventListener('click', () => { const v2=(this._q('#cli') as HTMLTextAreaElement)?.value.trim(); if(!v2)return; this._addLog('company',this.aC!,v2); this._renderCD(); });
    this._qa('[data-cv]').forEach(el => el.addEventListener('click', () => this._openCell(el.getAttribute('data-cv')!, el.getAttribute('data-cc')!)));
  }

  private _renderPD(): void {
    const mc = this._q('#main')!;
    this._ensureMx();
    const vids = Object.keys(this.V).filter(v => this.V[v].status!=='parsing');
    const cids = Object.keys(this.C);
    const pursue: any[]=[], cond: any[]=[], nogo: any[]=[];
    vids.forEach(vid => cids.forEach(cid => {
      const pw = this.M[vid][cid].po!=null?this.M[vid][cid].po:this._cpwin(vid,cid);
      const vs = this.M[vid][cid].so!=null?this.M[vid][cid].so:this._cscore(vid,cid);
      const item = {vid,cid,score:pw,vs,pursue:this.M[vid][cid].pursue,stage:this.M[vid][cid].stage};
      if (this.M[vid][cid].pursue==='pursue') pursue.push(item);
      else if (this.M[vid][cid].pursue==='conditional') cond.push(item);
      else nogo.push(item);
    }));
    const rf = (item: any) => { const v2=this.V[item.vid],co=this.C[item.cid],c=this._sc(item.score),st=item.stage||'Tracking'; return `<tr><td>${co.icon} <b>${co.name}</b></td><td style="font-weight:500">${v2.dn}</td><td><span style="font-size:14px;font-weight:700;color:${c}">${item.score}%</span></td><td style="font-size:11px">${item.vs!=null?item.vs.toLocaleString()+' pts':'-'}</td><td><span style="font-size:10px;font-weight:700;padding:2px 6px;border-radius:8px;background:${this._sbg(st)};color:${this._sfg(st)}">${st}</span></td><td><button class="btn bg" style="padding:3px 8px;font-size:10px" data-cv="${item.vid}" data-cc="${item.cid}">Edit</button></td></tr>`; };
    const thead = '<thead><tr><th>Company</th><th>Vehicle</th><th>P-Win</th><th>Score</th><th>Stage</th><th></th></tr></thead>';
    mc.innerHTML = `<div class="hrow"><div class="htitle">Pursuit Decisions</div><div style="display:flex;gap:10px;font-size:12px"><span style="color:#107c10;font-weight:600">${pursue.length} Pursue</span><span style="color:#835c00;font-weight:600">${cond.length} Conditional</span><span style="color:#a4262c;font-weight:600">${nogo.length} No-Go</span></div></div>${pursue.length?`<div class="stitle" style="color:#107c10">Pursue</div><table class="t" style="margin-bottom:14px">${thead}<tbody>${pursue.sort((a,b)=>b.score-a.score).map(rf).join('')}</tbody></table>`:''} ${cond.length?`<div class="stitle" style="color:#835c00">Conditional</div><table class="t" style="margin-bottom:14px">${thead}<tbody>${cond.sort((a,b)=>b.score-a.score).map(rf).join('')}</tbody></table>`:''} ${nogo.length?`<div class="stitle" style="color:#a4262c">No-Go</div><table class="t">${thead}<tbody>${nogo.sort((a,b)=>b.score-a.score).map(rf).join('')}</tbody></table>`:''} ${!pursue.length&&!cond.length&&!nogo.length?'<div class="empty"><h2>No Decisions Yet</h2><p>Click any matrix cell to set pursuit decisions.</p></div>':''}`;
    this._qa('[data-cv]').forEach(el => el.addEventListener('click', () => this._openCell(el.getAttribute('data-cv')!, el.getAttribute('data-cc')!)));
  }

  private _renderTL(): void {
    const mc = this._q('#main')!;
    const vids = Object.keys(this.V).filter(v => this.V[v].status==='parsed'||this.V[v].status==='raw');
    if (!vids.length) { mc.innerHTML='<div class="empty"><h2>No Vehicles</h2><p>Upload RFPs to see the timeline.</p></div>'; return; }
    const now=new Date(), start=new Date(now), end=new Date(now);
    start.setMonth(start.getMonth()-2); end.setMonth(end.getMonth()+12);
    const tot = end.getTime()-start.getTime();
    const pct = (ds: string) => { if (!ds||ds==='TBD') return null; const d=new Date(ds); if (isNaN(d.getTime())) return null; return Math.max(0,Math.min(100,((d.getTime()-start.getTime())/tot)*100)); };
    const mc2: any={releaseDate:'#1b3a6b',dueDate:'#835c00',awardDate:'#107c10'};
    const ml: any={releaseDate:'Release',dueDate:'Proposals Due',awardDate:'Expected Award'};
    const tp = ((now.getTime()-start.getTime())/tot)*100;
    const rows = vids.map(vid => {
      const v=this.V[vid],p=v.parsed;
      const dates: any={releaseDate:(v.de?.releaseDate)||(p?.releaseDate)||'TBD',dueDate:(v.de?.dueDate)||(p?.dueDate)||'TBD',awardDate:(v.de?.awardDate)||(p?.awardDate)||'TBD'};
      const ms = Object.keys(dates).map(key => { const x=pct(dates[key]); if(x===null) return `<div style="position:absolute;right:6px;top:50%;transform:translateY(-50%);font-size:9px;color:#a19f9d">${ml[key]}: TBD <button style="font-size:9px;padding:1px 4px;border:1px solid #c8c6c4;border-radius:3px;background:#fff;cursor:pointer" data-pd="${key}" data-pv="${vid}" data-pl="${ml[key]}">Set</button></div>`; return `<div style="position:absolute;top:50%;left:${x}%;transform:translate(-50%,-50%);cursor:pointer"><div style="width:12px;height:12px;border-radius:50%;border:2px solid #fff;background:${mc2[key]}"></div><div style="position:absolute;bottom:18px;left:50%;transform:translateX(-50%);background:#323130;color:#fff;font-size:10px;padding:3px 7px;border-radius:4px;white-space:nowrap;pointer-events:none;opacity:0" class="tltip">${ml[key]}: ${dates[key]}</div></div>`; }).join('');
      return `<div style="display:flex;align-items:center;margin-bottom:10px;gap:8px"><div style="width:130px;font-size:11px;font-weight:600;flex-shrink:0;text-align:right;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${v.dn}">${v.dn}</div><div style="flex:1;height:32px;position:relative;background:#f3f2f1;border-radius:4px;min-width:400px"><div style="position:absolute;top:0;bottom:0;width:2px;background:#a4262c;opacity:.5;left:${tp}%"></div>${ms}</div></div>`;
    }).join('');
    const axis: string[]=[];
    for (let m2=0;m2<=14;m2+=2){const d2=new Date(start);d2.setMonth(d2.getMonth()+m2);axis.push(`<div style="flex:1;text-align:center;font-size:10px;color:#605e5c">${d2.toLocaleDateString('en-US',{month:'short',year:'2-digit'})}</div>`);}
    mc.innerHTML = `<div class="hrow"><div class="htitle">Opportunity Timeline</div></div><div style="display:flex;gap:14px;margin-bottom:10px;font-size:11px;flex-wrap:wrap"><span>Navy=Release</span><span>Amber=Due</span><span>Green=Award</span><span>Red=Today</span></div><div style="overflow-x:auto"><div style="display:flex;margin-left:138px;min-width:400px">${axis.join('')}</div>${rows}</div>`;
    this._qa('[data-pd]').forEach(el => el.addEventListener('click', () => { this.dtCb={vid:el.getAttribute('data-pv'),field:el.getAttribute('data-pd')}; (this._q('#dtTitle') as HTMLElement).textContent='Set Date - '+el.getAttribute('data-pl'); (this._q('#dtInput') as HTMLInputElement).value=''; this._q('#dtModal')?.classList.add('on'); }));
  }

  private _saveDt(): void {
    if (!this.dtCb) return;
    const val = (this._q('#dtInput') as HTMLInputElement).value;
    if (val) { if (!this.V[this.dtCb.vid].de) this.V[this.dtCb.vid].de={}; this.V[this.dtCb.vid].de[this.dtCb.field]=val; this._saveV(this.V[this.dtCb.vid]); }
    this._q('#dtModal')?.classList.remove('on'); this.dtCb=null; this._renderContent();
  }

  private _renderIF(): void {
    const mc = this._q('#main')!;
    const all = [...this.logs].reverse();
    mc.innerHTML = `<div class="hrow"><div class="htitle">Intel Feed <span style="font-size:12px;font-weight:400;color:#605e5c">${all.length} entries</span></div></div><div>${all.length?all.map(l=>`<div class="log" style="border-left-color:${l.author==='Deuce'?'#1b3a6b':'#edebe9'}"><div class="logm">${l.author} - ${l.scope} - ${l.date}</div>${l.entry}</div>`).join(''):'<div class="empty" style="padding:30px"><p>No intel entries yet.</p></div>'}</div>`;
  }

  private _saveComp(): void {
    const nm = (this._q('#svn') as HTMLInputElement).value.trim(); if (!nm){alert('Enter a name.');return;}
    const comp = {id:'s'+Date.now(),name:nm,notes:(this._q('#svnt') as HTMLTextAreaElement).value.trim(),date:new Date().toLocaleDateString(),author:this.context.pageContext.user.displayName,companies:JSON.parse(JSON.stringify(this.C)),matrix:JSON.parse(JSON.stringify(this.M)),pw:JSON.parse(JSON.stringify(this.PW))};
    this.comps.unshift(comp); this._saveCompSP(comp);
    this._q('#svModal')?.classList.remove('on');
    this._addLog('system','platform','Comparison saved: '+nm);
    this._renderContent();
  }

  private _renderSV(): void {
    const mc = this._q('#main')!;
    mc.innerHTML = `<div class="hrow"><div class="htitle">Saved Comparisons <span style="font-size:12px;font-weight:400;color:#605e5c">${this.comps.length} saved</span></div><button class="btn bp" id="svobtn2">Save Current</button></div>${this.comps.length?this.comps.map(s=>`<div class="card"><div style="padding:12px 14px;display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:8px"><div><div style="font-size:13px;font-weight:600">${s.name}</div><div style="font-size:11px;color:#605e5c;margin-top:2px">${s.date} - ${s.author} - ${Object.keys(s.companies||{}).length} companies</div>${s.notes?`<div style="font-size:11px;color:#605e5c;margin-top:3px">${s.notes}</div>`:''}</div><div style="display:flex;gap:6px"><button class="btn bg" data-lc="${s.id}">Load</button><button class="btn bd" data-dc2="${s.id}">Delete</button></div></div></div>`).join(''):'<div class="empty" style="padding:30px"><p>No saved comparisons yet.</p></div>'}`;
    this._q('#svobtn2')?.addEventListener('click', () => this._q('#svModal')?.classList.add('on'));
    this._qa('[data-lc]').forEach(el => el.addEventListener('click', () => { const s=this.comps.find(c=>c.id===el.getAttribute('data-lc')); if(!s)return; if(!confirm(`Load "${s.name}"?`))return; Object.keys(s.matrix||{}).forEach(k=>this.M[k]=s.matrix[k]); Object.keys(s.pw||{}).forEach(k=>this.PW[k]=s.pw[k]); Object.keys(s.companies||{}).forEach(cid=>{if(!this.C[cid])this.C[cid]=s.companies[cid];}); this._addLog('system','platform','Loaded: '+s.name); this._setNav(0); this._renderContent(); }));
    this._qa('[data-dc2]').forEach(el => el.addEventListener('click', () => { const idx=this.comps.findIndex(c=>c.id===el.getAttribute('data-dc2')); if(idx>-1)this.comps.splice(idx,1); this._renderSV(); }));
  }

  // ── COMPANY MODAL ─────────────────────────────────────────────────────────────
  private _openCoModal(eid: string|null): void {
    (this._q('#coTitle') as HTMLElement).textContent = eid?'Edit Company':'Add Company';
    (this._q('#eid') as HTMLInputElement).value = eid||'';
    const co = eid?this.C[eid]:{};
    ['cn','cs','cc','cu','chq','cap','rev'].forEach(id => (this._q('#'+id) as HTMLInputElement).value = co[id==='cn'?'name':id==='cs'?'short':id==='cc'?'cage':id==='cu'?'uei':id==='chq'?'hq':id]||'');
    const cweb = this._q('#cweb') as HTMLInputElement; if(cweb) cweb.value = co.web||'';
    (this._q('#cico') as HTMLInputElement).value = co.icon||'*';
    (this._q('#nts') as HTMLTextAreaElement).value = co.notes||'';
    (this._q('#csz') as HTMLSelectElement).value = co.size||'Small Business';
    (this._q('#ccl') as HTMLSelectElement).value = co.clearance||'None';
    (this._q('#cmmc') as HTMLSelectElement).value = co.cmmc||'0';
    (this._q('#cna') as HTMLInputElement).value = (co.naics||[]).join(', ');
    (this._q('#cps') as HTMLInputElement).value = (co.psc||[]).join(', ');
    (this._q('#dnote') as HTMLElement).style.display = 'none';
    this._bldCkg('btg', this.BT, co.bizTypes||[]);
    this._bldCkg('ceg', this.CE, co.certs||[]);
    this._bldCkg('dcg', this.DC, co.dcma||[]);
    this._qa('.itab').forEach((t,i) => { t.classList.toggle('on',i===0); (t as HTMLElement).style.background=i===0?'#1b3a6b':'transparent'; (t as HTMLElement).style.color=i===0?'#fff':'#605e5c'; });
    this._qa('.ipane').forEach((p,i) => p.classList.toggle('on',i===0));
    this._q('#coModal')?.classList.add('on');
  }
  private _closeCoModal(): void { this._q('#coModal')?.classList.remove('on'); }

  private _bldCkg(id: string, items: string[], sel: string[]): void {
    const el = this._q('#'+id)!;
    el.innerHTML = items.map(item => `<label class="cki${sel.includes(item)?' on':''}"><input type="checkbox" value="${item}"${sel.includes(item)?' checked':''}>${item}</label>`).join('');
    el.querySelectorAll('.cki').forEach(lab => lab.querySelector('input')?.addEventListener('change', function(){ lab.classList.toggle('on', (this as HTMLInputElement).checked); }));
  }
  private _gCkg(id: string): string[] { return Array.from(this.domElement.querySelectorAll('#'+id+' input:checked')).map((i: any) => i.value); }

  private _saveCompany(): void {
    const nm = (this._q('#cn') as HTMLInputElement).value.trim(); if (!nm){alert('Company name required.');return;}
    const eid = (this._q('#eid') as HTMLInputElement).value;
    const id = eid||('c'+this._uid());
    this.C[id] = { id, name:nm, short:(this._q('#cs') as HTMLInputElement).value.trim()||nm.split(' ')[0], cage:(this._q('#cc') as HTMLInputElement).value.trim(), web:(this._q('#cweb') as HTMLInputElement)?.value.trim(), uei:(this._q('#cu') as HTMLInputElement).value.trim(), icon:(this._q('#cico') as HTMLInputElement).value||'*', hq:(this._q('#chq') as HTMLInputElement).value.trim(), size:(this._q('#csz') as HTMLSelectElement).value, bizTypes:this._gCkg('btg'), naics:(this._q('#cna') as HTMLInputElement).value.split(',').map(s=>s.trim()).filter(Boolean), psc:(this._q('#cps') as HTMLInputElement).value.split(',').map(s=>s.trim()).filter(Boolean), clearance:(this._q('#ccl') as HTMLSelectElement).value, certs:this._gCkg('ceg'), cmmc:(this._q('#cmmc') as HTMLSelectElement).value, dcma:this._gCkg('dcg'), caps:(this._q('#cap') as HTMLInputElement).value.trim(), rev:(this._q('#rev') as HTMLInputElement).value.trim(), notes:(this._q('#nts') as HTMLTextAreaElement).value.trim() };
    if (!eid) this._addLog('company', id, 'Profile created: '+nm, 'system');
    this._saveCo(this.C[id]);
    this._closeCoModal(); this._ensureMx(); this._renderSidebar(); this._renderContent();
  }

  private _popCoForm(p: any): void {
    (this._q('#cn') as HTMLInputElement).value = p.name||'';
    (this._q('#cs') as HTMLInputElement).value = p.short||'';
    (this._q('#cc') as HTMLInputElement).value = p.cage||'';
    (this._q('#cu') as HTMLInputElement).value = p.uei||'';
    (this._q('#chq') as HTMLInputElement).value = p.hq||'';
    (this._q('#rev') as HTMLInputElement).value = p.rev||'';
    const cwebf = this._q('#cweb') as HTMLInputElement; if(cwebf) cwebf.value = p.web||'';
    (this._q('#cap') as HTMLInputElement).value = p.caps||'';
    (this._q('#nts') as HTMLTextAreaElement).value = p.notes||'';
    if (p.size) (this._q('#csz') as HTMLSelectElement).value = p.size;
    if (p.clearance) (this._q('#ccl') as HTMLSelectElement).value = p.clearance;
    if (p.cmmc) (this._q('#cmmc') as HTMLSelectElement).value = p.cmmc;
    (this._q('#cna') as HTMLInputElement).value = (p.naics||[]).join(', ');
    (this._q('#cps') as HTMLInputElement).value = (p.psc||[]).join(', ');
    this._bldCkg('btg', this.BT, p.bizTypes||[]);
    this._bldCkg('ceg', this.CE, p.certs||[]);
    this._bldCkg('dcg', this.DC, p.dcma||[]);
    if (p.sourceNote) { const dn=this._q('#dnote') as HTMLElement; dn.textContent='Deuce: '+p.sourceNote; dn.style.display='block'; }
  }

  private _parseFromPaste(): void {
    const txt = (this._q('#ptxt') as HTMLTextAreaElement).value.trim();
    const nm  = (this._q('#pcn') as HTMLInputElement).value.trim();
    if (!txt){alert('Please paste some text.');return;}
    this._ps('pstat','wait','Deuce is analyzing...');
    this._callAPI(this.COP+(nm?'Company focus: '+nm+'\n\n':'')+txt.slice(0,12000), (err,p) => {
      if (err||!p){this._ps('pstat','err','Could not parse. Error: '+(err?.message||'invalid response'));return;}
      this._popCoForm(p); this._ps('pstat','ok','Extracted: '+p.name+'. Review fields then Save.');
      this._qa('.itab').forEach((t,i) => { t.classList.toggle('on',i===0); (t as HTMLElement).style.background=i===0?'#1b3a6b':'transparent'; (t as HTMLElement).style.color=i===0?'#fff':'#605e5c'; });
      this._qa('.ipane').forEach((p2,i) => p2.classList.toggle('on',i===0));
    });
  }

  private _parseFromUpload(): void {
    if (!this.uFile){alert('Please select a file.');return;}
    this._ps('ustat','wait','Deuce is reading...');
    const r = new FileReader();
    r.onload = (e) => {
      const txt = typeof e.target?.result==='string' ? e.target.result : this._extractPdf(e.target?.result as ArrayBuffer);
      if (!txt||txt.length<50){this._ps('ustat','err','Could not extract text. Use Paste Text tab.');return;}
      const nm = (this._q('#ucn') as HTMLInputElement).value.trim();
      this._callAPI(this.COP+(nm?'Company focus: '+nm+'\n\n':'')+txt.slice(0,12000), (err,p) => {
        if (err||!p){this._ps('ustat','err','Could not parse. Error: '+(err?.message||'invalid response'));return;}
        this._popCoForm(p); this._ps('ustat','ok','Extracted: '+p.name+'. Review fields then Save.');
        this._qa('.itab').forEach((t,i) => { t.classList.toggle('on',i===0); (t as HTMLElement).style.background=i===0?'#1b3a6b':'transparent'; (t as HTMLElement).style.color=i===0?'#fff':'#605e5c'; });
        this._qa('.ipane').forEach((p2,i) => p2.classList.toggle('on',i===0));
      });
    };
    if (this.uFile.name.toLowerCase().includes('.pdf')) r.readAsArrayBuffer(this.uFile);
    else r.readAsText(this.uFile);
  }

  private _extractPdf(buf: ArrayBuffer): string {
    const b = new Uint8Array(buf); let s='';
    for (let i=0;i<b.length;i++) s+=String.fromCharCode(b[i]);
    const m = s.match(/\(([^)]{3,200})\)/g)||[];
    return m.slice(0,400).map(x=>x.slice(1,-1)).filter(x=>/[a-zA-Z]{3,}/.test(x)).join(' ');
  }

  private _handleRfpFiles(files: FileList|null|undefined): void {
    if (!files) return;
    Array.from(files).forEach(f => {
      const id = 'v'+this._uid();
      this.V[id] = {id,name:f.name,size:f.size,status:'parsing',raw:'',parsed:null,dn:f.name.replace(/\.[^.]+$/,''),de:{}};
      this._renderSidebar();
      const r = new FileReader();
      r.onload = (e) => {
        const txt = typeof e.target?.result==='string' ? e.target.result : this._extractPdf(e.target?.result as ArrayBuffer);
        this.V[id].raw = txt.slice(0,8000);
        this._callAPI(this.RFPP+txt.slice(0,12000), (err,p) => {
          if (err||!p){this.V[id].parsed=this._fallback(this.V[id]);this.V[id].status='raw';}
          else{this.V[id].parsed=p;this.V[id].dn=p.vehicleName||this.V[id].dn;this.V[id].status='parsed';this._addLog('vehicle',id,'RFP parsed: '+p.vehicleName,'Deuce');}
          this._saveV(this.V[id]);this._ensureMx();this._renderSidebar();if(this.nav===0||this.nav===1)this._renderContent();
        });
      };
      if (f.name.toLowerCase().includes('.pdf')) r.readAsArrayBuffer(f); else r.readAsText(f);
    });
  }

  private _handleCoFile(file: File|null|undefined): void {
    if (!file) return;
    this.uFile = file;
    const fn = this._q('#ufn') as HTMLElement;
    fn.style.display='block'; fn.textContent=file.name+' ('+Math.round(file.size/1024)+'KB)';
  }

  private _submitRfp(): void {
    const txt = (this._q('#rt') as HTMLTextAreaElement).value.trim();
    const nm  = (this._q('#rn') as HTMLInputElement).value.trim()||'Unnamed RFP';
    if (!txt){alert('Please paste RFP text.');return;}
    this._q('#rpModal')?.classList.remove('on');
    const id = 'v'+this._uid();
    this.V[id]={id,name:nm+'.txt',size:txt.length,status:'parsing',raw:txt.slice(0,8000),parsed:null,dn:nm,de:{}};
    this._renderSidebar();
    this._callAPI(this.RFPP+txt.slice(0,12000), (err,p) => {
      if (err||!p){this.V[id].parsed=this._fallback(this.V[id]);this.V[id].status='raw';}
      else{this.V[id].parsed=p;this.V[id].dn=p.vehicleName||this.V[id].dn;this.V[id].status='parsed';this._addLog('vehicle',id,'RFP parsed: '+p.vehicleName,'Deuce');}
      this._saveV(this.V[id]);this._ensureMx();this._renderSidebar();if(this.nav===0||this.nav===1)this._renderContent();
    });
  }

  private _fallback(v: any): any {
    return {vehicleName:v.dn,agency:'Review RFP',type:'Unknown',pop:'See RFP',naics:[],ceiling:'TBD',releaseDate:'TBD',dueDate:'TBD',awardDate:'TBD',solNum:'TBD',sbSlots:'Review RFP',hasScorecard:false,scorecardSections:[],gates:[{num:'01',label:'Manual Review Required',color:'#1b3a6b',rows:[['Review requirements','Could not auto-parse','VERIFY','vf']]}],domains:[{name:'Primary Domain',naics:'TBD',pct:50,rec:true,color:'#1b3a6b',note:'Update after review',caps:[['Update based on RFP','par']]}],pwinFactors:[{key:'rel',label:'Customer Intimacy',desc:'Agency relationships',low:'None',mid:'Indirect',high:'Direct',color:'#1b3a6b',val:1},{key:'pp',label:'Past Performance',desc:'Relevant contracts',low:'None',mid:'Related',high:'Direct',color:'#107c10',val:1},{key:'comp',label:'Competitive Position',desc:'vs. field',low:'Weak',mid:'Competitive',high:'Strong',color:'#5c2d91',val:1},{key:'sol',label:'Solution Maturity',desc:'Readiness',low:'Concept',mid:'Partial',high:'Proven',color:'#835c00',val:1},{key:'team',label:'Teaming Strength',desc:'Partners',low:'None',mid:'Potential',high:'Established',color:'#d83b01',val:1}],scoreCalcRules:[],scoreRef:[['Past Performance','TBD','Review RFP'],['Technical','TBD','Review RFP'],['Price','TBD','Review RFP']],defaultDecision:'conditional',decisionRationale:'Could not auto-parse. Use Paste RFP option.'};
  }

  private _scrapeWebsite(): void {
    const urlEl = this._q('#cweb') as HTMLInputElement;
    const url = urlEl?.value.trim();
    if (!url){alert('Enter a website URL first.');return;}
    this._ps('scrapestat','wait','Deuce is scraping '+url+'...');
    fetch('https://twcg-proxy.kyle-88e.workers.dev/fetch', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({url})
    }).then(r=>r.json()).then(data => {
      if (data.error){this._ps('scrapestat','err','Scrape failed: '+data.error);return;}
      const txt = data.text||'';
      this._ps('scrapestat','wait','Deuce is analyzing '+data.pagesScraped+' pages...');
      this._callAPI(this.COP+'Website: '+url+'\n\n'+txt.slice(0,12000), (err,p) => {
        if (err||!p){this._ps('scrapestat','err','Could not parse. Error: '+(err?.message||'invalid response'));return;}
        this._popCoForm({...p, web:url});
        this._ps('scrapestat','ok','Scraped '+data.pagesScraped+' pages. '+data.linksFound+' links found. Review fields then Save.');
      });
    }).catch(e => this._ps('scrapestat','err','Fetch error: '+e.message));
  }

  private enrichFile: File | null = null;
  private enrichTargetId: string | null = null;
  private enrichParsed: any = null;

  private _parseFromWebsite(): void {
    const txt = (this._q('#wtxt') as HTMLTextAreaElement).value.trim();
    const nm = (this._q('#wcn') as HTMLInputElement).value.trim();
    if (!txt){alert('Please paste website text.');return;}
    this._ps('wstat','wait','Deuce is analyzing...');
    this._callAPI(this.COP+(nm?'Company focus: '+nm+'\n':'')+txt.slice(0,12000), (err,p) => {
      if (err||!p){this._ps('wstat','err','Could not parse. Error: '+(err?.message||'invalid response'));return;}
      this._popCoForm(p); this._ps('wstat','ok','Extracted: '+p.name+'. Review fields then Save.');
      this._qa('.itab').forEach((t,i) => { t.classList.toggle('on',i===0); (t as HTMLElement).style.background=i===0?'#1b3a6b':'transparent'; (t as HTMLElement).style.color=i===0?'#fff':'#605e5c'; });
      this._qa('.ipane').forEach((p2,i) => p2.classList.toggle('on',i===0));
    });
  }

  private _openEnrichModal(cid: string): void {
    this.enrichTargetId = cid;
    this.enrichFile = null;
    this.enrichParsed = null;
    (this._q('#eptxt') as HTMLTextAreaElement).value = '';
    (this._q('#ewtxt') as HTMLTextAreaElement).value = '';
    const fn = this._q('#ecfn') as HTMLElement; if(fn) fn.style.display='none';
    this._qa('[data-et]').forEach((t,i) => { t.classList.toggle('on',i===0); (t as HTMLElement).style.background=i===0?'#1b3a6b':'transparent'; (t as HTMLElement).style.color=i===0?'#fff':'#605e5c'; });
    this._qa('.epane').forEach((p,i) => p.classList.toggle('on',i===0));
    ['epstat','ecstat','ewstat'].forEach(id => { const el=this._q('#'+id); if(el){el.className='pstat';el.textContent='';} });
    this._q('#enrichSrcModal')?.classList.add('on');
  }

  private _enrichFromPaste(): void {
    const txt = (this._q('#eptxt') as HTMLTextAreaElement).value.trim();
    if (!txt){alert('Please paste some text.');return;}
    this._ps('epstat','wait','Deuce is analyzing...');
    this._callAPI(this.COP+txt.slice(0,12000), (err,p) => {
      if (err||!p){this._ps('epstat','err','Could not parse.');return;}
      this._ps('epstat','ok','Done. Review changes.');
      this._showEnrichDiff(p);
    });
  }

  private _enrichFromUpload(): void {
    if (!this.enrichFile){alert('Please select a file.');return;}
    this._ps('ecstat','wait','Deuce is reading...');
    const r = new FileReader();
    r.onload = (e) => {
      const txt = typeof e.target?.result==='string' ? e.target.result : this._extractPdf(e.target?.result as ArrayBuffer);
      if (!txt||txt.length<50){this._ps('ecstat','err','Could not extract text.');return;}
      this._callAPI(this.COP+txt.slice(0,12000), (err,p) => {
        if (err||!p){this._ps('ecstat','err','Could not parse.');return;}
        this._ps('ecstat','ok','Done. Review changes.');
        this._showEnrichDiff(p);
      });
    };
    if (this.enrichFile.name.toLowerCase().includes('.pdf')) r.readAsArrayBuffer(this.enrichFile);
    else r.readAsText(this.enrichFile);
  }

  private _enrichFromWebsite(): void {
    const txt = (this._q('#ewtxt') as HTMLTextAreaElement).value.trim();
    if (!txt){alert('Please paste website text.');return;}
    this._ps('ewstat','wait','Deuce is analyzing...');
    this._callAPI(this.COP+txt.slice(0,12000), (err,p) => {
      if (err||!p){this._ps('ewstat','err','Could not parse.');return;}
      this._ps('ewstat','ok','Done. Review changes.');
      this._showEnrichDiff(p);
    });
  }

  private _showEnrichDiff(newData: any): void {
    if (!this.enrichTargetId) return;
    this.enrichParsed = newData;
    const cur = this.C[this.enrichTargetId];
    const fields = [
      {key:'name',label:'Company Name'},{key:'short',label:'Short Name'},
      {key:'cage',label:'CAGE'},{key:'uei',label:'UEI'},
      {key:'hq',label:'HQ'},{key:'rev',label:'Revenue'},
      {key:'caps',label:'Capabilities'},{key:'notes',label:'BD Notes'},
      {key:'clearance',label:'Clearance'},{key:'cmmc',label:'CMMC'},
      {key:'size',label:'Business Size'}
    ];
    const rows = fields.map(f => {
      const cv = (cur as any)[f.key]||''; const nv = newData[f.key]||'';
      if (!nv||nv===cv) return '';
      if (!cv) return `<div style="margin-bottom:8px;padding:8px;background:#dff6dd;border-radius:6px"><div style="font-size:10px;font-weight:700;color:#605e5c;margin-bottom:3px">${f.label}</div><div style="font-size:12px">Auto-fill: <b>${nv}</b></div><input type="hidden" class="enrich-auto" data-fk="${f.key}" data-fv="${nv}"></div>`;
      return `<div style="margin-bottom:8px;padding:8px;background:#fff4ce;border-radius:6px;border:1px solid #edebe9"><div style="font-size:10px;font-weight:700;color:#605e5c;margin-bottom:5px">${f.label} — choose one:</div><div style="display:flex;gap:6px;flex-wrap:wrap"><button class="btn bg enrich-pick on" data-fk="${f.key}" data-fv="${cv}" style="flex:1;text-align:left;font-size:11px;padding:5px 8px;border:2px solid #1b3a6b">KEEP: ${cv}</button><button class="btn bg enrich-pick" data-fk="${f.key}" data-fv="${nv}" style="flex:1;text-align:left;font-size:11px;padding:5px 8px">NEW: ${nv}</button></div></div>`;
    }).filter(Boolean).join('');
    const diff = this._q('#enrichDiff')!;
    diff.innerHTML = rows || '<div style="padding:10px;color:#107c10;font-weight:600">No conflicting fields — all new data will auto-fill empty fields.</div>';
    diff.querySelectorAll('.enrich-pick').forEach(btn => {
      btn.addEventListener('click', () => {
        const fk = btn.getAttribute('data-fk')!;
        const grp = diff.querySelectorAll(`.enrich-pick[data-fk="${fk}"]`);
        grp.forEach(b => { b.classList.remove('on'); (b as HTMLElement).style.border='1px solid #edebe9'; });
        btn.classList.add('on'); (btn as HTMLElement).style.border='2px solid #1b3a6b';
      });
    });
    this._q('#enrichSrcModal')?.classList.remove('on');
    this._q('#enrichModal')?.classList.add('on');
  }

  private _applyEnrich(): void {
    if (!this.enrichTargetId || !this.enrichParsed) return;
    const co = this.C[this.enrichTargetId];
    const diff = this._q('#enrichDiff')!;
    // Apply auto-fills
    diff.querySelectorAll('.enrich-auto').forEach(el => {
      const fk = el.getAttribute('data-fk')!; const fv = el.getAttribute('data-fv')!;
      (co as any)[fk] = fv;
    });
    // Apply picks
    const picked: any = {};
    diff.querySelectorAll('.enrich-pick.on').forEach(btn => {
      picked[btn.getAttribute('data-fk')!] = btn.getAttribute('data-fv')!;
    });
    Object.keys(picked).forEach(k => (co as any)[k] = picked[k]);
    // Append new naics/psc/bizTypes/certs/dcma
    ['naics','psc','bizTypes','certs','dcma'].forEach(k => {
      const existing: string[] = (co as any)[k] || [];
      const incoming: string[] = (this.enrichParsed as any)[k] || [];
      const merged = Array.from(new Set([...existing, ...incoming]));
      (co as any)[k] = merged;
    });
    this._addLog('company', this.enrichTargetId, 'Profile enriched from new source via Deuce');
    this._saveCo(co);
    this._q('#enrichModal')?.classList.remove('on');
    this.enrichTargetId = null; this.enrichParsed = null;
    this._renderSidebar(); this._renderContent();
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }
}