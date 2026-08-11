/** @jsxImportSource react */
import React, { useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import { Area, AreaChart, CartesianGrid, ResponsiveContainer, ReferenceLine, Tooltip, XAxis, YAxis } from 'recharts'
import { AlertTriangle, BarChart3, CheckCircle2, ChevronDown, FileSpreadsheet, Filter, Gauge, Layers3, UploadCloud, X, Zap } from 'lucide-react'

type PositionValues = { 전: number | null; 중: number | null; 후: number | null }
type QualityRow = { id:number; date:string; time:string; client:string; grade:string; basisWeight:number|null; source:string; values: Record<string, PositionValues>; averages: Record<string, number|null> }

type MetricDef = { label:string; key:string; unit:string; usl:number; lsl:number; kind:'position'|'single' }
const metricOptions: MetricDef[] = [
  { label:'평량 (g/㎡)', key:'평량', unit:'g/㎡', usl:306, lsl:294, kind:'position' },
  { label:'두께 (㎛)', key:'두께', unit:'㎛', usl:360, lsl:340, kind:'position' },
  { label:'수분 (%)', key:'수분', unit:'%', usl:7.5, lsl:4.5, kind:'position' },
  { label:'백색도', key:'백색도', unit:'', usl:80, lsl:42, kind:'position' },
  { label:'PPS', key:'PPS', unit:'', usl:3.5, lsl:1, kind:'position' },
  { label:'인쇄층분리', key:'인쇄층분리', unit:'', usl:100, lsl:15, kind:'single' },
  { label:'픽킹 (TV20)', key:'픽킹', unit:'', usl:100, lsl:0, kind:'single' },
  { label:'모틀링', key:'모틀링', unit:'', usl:100, lsl:3, kind:'single' },
  { label:'내절도', key:'내절도', unit:'회', usl:100, lsl:0, kind:'position' },
  { label:'스티프니스', key:'스티프니스', unit:'', usl:150, lsl:0, kind:'position' },
  { label:'파열강도', key:'파열강도', unit:'', usl:150, lsl:0, kind:'single' },
  { label:'인터널', key:'인터널', unit:'', usl:100, lsl:0, kind:'single' },
]
const positionKeys = ['전','중','후'] as const
const emptyPositions = ():PositionValues => ({ 전:null, 중:null, 후:null })
const numeric = (value:unknown):number|null => { const n=Number(value); return Number.isFinite(n) ? n : null }
const slashAverage = (value:unknown):number|null => { const nums=String(value??'').split('/').map((v)=>Number(v.trim())).filter(Number.isFinite); return nums.length ? nums.reduce((a,b)=>a+b,0)/nums.length : numeric(value) }
const pos = (row:(string|number|null)[], indices:number[]):PositionValues => ({ 전:numeric(row[indices[0]]), 중:numeric(row[indices[1]]), 후:numeric(row[indices[2]]) })
const avgPosition = (values:PositionValues):number|null => { const nums=positionKeys.map(k=>values[k]).filter((v):v is number=>v!==null); return nums.length ? nums.reduce((a,b)=>a+b,0)/nums.length : null }
const metricValue = (values:PositionValues):number|null => avgPosition(values)

function rowFromSheet(row:(string|number|null)[], index:number):QualityRow|null {
  const client=String(row[5]??'').replace(/\n/g,' ').trim(); const grade=String(row[6]??'').trim(); const basisWeight=numeric(row[7]);
  const hasData=client || grade || row.slice(8).some((v)=>v!==null && v!==''); if(!hasData) return null
  const values:Record<string,PositionValues> = {
    평량:pos(row,[8,9,10]), 두께:pos(row,[11,12,13]), 수분:pos(row,[15,16,17]),
    백색도:{ 전:avgPosition(pos(row,[18,19,20])), 중:null, 후:numeric(row[21]) },
    PPS:{ 전:avgPosition(pos(row,[23,24,25])), 중:null, 후:numeric(row[26]) },
    인쇄층분리:{ 전:slashAverage(row[30]), 중:null, 후:null }, 픽킹:{ 전:numeric(row[31]), 중:null, 후:null },
    모틀링:{ 전:numeric(row[32]), 중:null, 후:null }, 내절도:{ 전:slashAverage(row[37]), 중:slashAverage(row[38]), 후:null },
    스티프니스:{ 전:numeric(row[39]), 중:numeric(row[40]), 후:null }, 파열강도:{ 전:numeric(row[41]), 중:null, 후:null },
    인터널:{ 전:slashAverage(row[45]), 중:null, 후:null },
  }
  const averages=Object.fromEntries(Object.entries(values).map(([key,v])=>[key,metricValue(v)]))
  const date=`${String(row[1]??'').padStart(2,'0')}.${String(row[2]??'').padStart(2,'0')}`
  const time=`${String(row[3]??'').padStart(2,'0')}:${String(row[4]??'').padStart(2,'0')}`
  return { id:index+1,date,time,client:client||'미지정 거래처',grade:grade||'미지정',basisWeight,source:client.slice(0,18)||`ROW-${index+1}`,values,averages }
}
function parseWorkbook(file:File):Promise<QualityRow[]> { return file.arrayBuffer().then(buffer=>{ const wb=XLSX.read(buffer,{type:'array',cellDates:true}); const name=wb.SheetNames.find(n=>n.includes('변경일지'))??wb.SheetNames[0]; const rows=XLSX.utils.sheet_to_json<(string|number|null)[]>(wb.Sheets[name],{header:1,defval:null}); return rows.slice(8).map(rowFromSheet).filter((r):r is QualityRow=>Boolean(r)) }) }

const demoRows:QualityRow[] = [
  {id:1,date:'08.10',time:'07:38',client:'중국 · 제약케이스',grade:'IV',basisWeight:300,source:'6810J30001',values:{평량:{전:303,중:301,후:302},두께:{전:349,중:344,후:349},수분:{전:6,중:6,후:6.1},백색도:{전:76.6,중:null,후:67.2},PPS:{전:2.6,중:null,후:6.7},인쇄층분리:{전:10,중:null,후:null},픽킹:{전:11,중:null,후:null},모틀링:{전:2.9,중:null,후:null},내절도:{전:19.5,중:20,후:null},스티프니스:{전:100,중:40,후:null},파열강도:{전:4.9,중:null,후:null},인터널:{전:.83,중:null,후:null}},averages:{}},
  {id:2,date:'08.10',time:'07:49',client:'신승아이엔씨',grade:'IV',basisWeight:300,source:'6810J30005',values:{평량:{전:305,중:304,후:305},두께:{전:357,중:355,후:358},수분:{전:6.3,중:6.3,후:6.2},백색도:{전:77,중:null,후:67.5},PPS:{전:2.6,중:null,후:null},인쇄층분리:{전:13,중:null,후:null},픽킹:{전:10,중:null,후:null},모틀링:{전:2.7,중:null,후:null},내절도:{전:18,중:20,후:null},스티프니스:{전:100,중:40,후:null},파열강도:{전:4.8,중:null,후:null},인터널:{전:.8,중:null,후:null}},averages:{}},
  {id:3,date:'08.10',time:'08:16',client:'중국 · 제약케이스',grade:'IV',basisWeight:300,source:'6810J30012',values:{평량:{전:307,중:306,후:308},두께:{전:362,중:361,후:363},수분:{전:7.9,중:7.8,후:7.7},백색도:{전:74,중:null,후:66},PPS:{전:3.8,중:null,후:7},인쇄층분리:{전:18,중:null,후:null},픽킹:{전:12,중:null,후:null},모틀링:{전:3.1,중:null,후:null},내절도:{전:14,중:15,후:null},스티프니스:{전:102,중:42,후:null},파열강도:{전:4.1,중:null,후:null},인터널:{전:1.2,중:null,후:null}},averages:{}},
  {id:4,date:'08.10',time:'08:32',client:'태국 · 포장용지',grade:'III',basisWeight:250,source:'6810J30018',values:{평량:{전:298,중:299,후:298},두께:{전:348,중:347,후:349},수분:{전:5.8,중:5.9,후:5.8},백색도:{전:75,중:null,후:65},PPS:{전:2.8,중:null,후:6.4},인쇄층분리:{전:9,중:null,후:null},픽킹:{전:8,중:null,후:null},모틀링:{전:2.4,중:null,후:null},내절도:{전:22,중:24,후:null},스티프니스:{전:95,중:38,후:null},파열강도:{전:5.2,중:null,후:null},인터널:{전:.7,중:null,후:null}},averages:{}},
]
for (const row of demoRows) row.averages=Object.fromEntries(Object.entries(row.values).map(([k,v])=>[k,metricValue(v)]))

function locationAverage(rows:QualityRow[], metric:string, location:'전'|'중'|'후'):number|null {
 const nums=rows.map(r=>r.values[metric]?.[location]).filter((v):v is number=>v!==null && v!==undefined)
 return nums.length ? nums.reduce((a,b)=>a+b,0)/nums.length : null
}

function App(){
 const inputRef=useRef<HTMLInputElement>(null); const [rows,setRows]=useState<QualityRow[]>(demoRows); const [fileName,setFileName]=useState('3호기 2026.08.10.xls'); const [grade,setGrade]=useState('전체 지종'); const [basis,setBasis]=useState('전체 평량'); const [metric,setMetric]=useState('평량'); const [usl,setUsl]=useState('306'); const [lsl,setLsl]=useState('294'); const [dragging,setDragging]=useState(false)
 const selected=metricOptions.find(m=>m.key===metric)??metricOptions[0]; const grades=['전체 지종',...Array.from(new Set(rows.map(r=>r.grade)))]; const bases=['전체 평량',...Array.from(new Set(rows.map(r=>r.basisWeight).filter((v):v is number=>v!==null))).sort((a,b)=>a-b).map(String)]
 const filtered=useMemo(()=>rows.filter(r=>(grade==='전체 지종'||r.grade===grade)&&(basis==='전체 평량'||String(r.basisWeight)===basis)),[rows,grade,basis]); const values=filtered.map(r=>r.averages[metric]).filter((v):v is number=>v!==null); const average=values.length?values.reduce((a,b)=>a+b,0)/values.length:0; const outliers=filtered.filter(r=>{const v=r.averages[metric];return v!==null&&(v>Number(usl)||v<Number(lsl))}); const chartData=filtered.map(r=>({label:`${r.date} ${r.time}`,value:r.averages[metric],row:r})).filter(d=>d.value!==null); const passRate=filtered.length?Math.round((filtered.length-outliers.length)/filtered.length*100):0
 const loadFile=async(file?:File)=>{if(!file||!/\.xls[x]?$/.test(file.name.toLowerCase()))return;const parsed=await parseWorkbook(file);setRows(parsed.length?parsed:demoRows);setFileName(file.name)}; const setMetricSafe=(key:string)=>{const m=metricOptions.find(x=>x.key===key)??metricOptions[0];setMetric(key);setUsl(String(m.usl));setLsl(String(m.lsl))}
 return <main className="app-shell"><header className="topbar"><div className="brand"><span className="brand-mark"><Zap size={16} fill="currentColor"/></span><span>QUALITY <b>LEDGER</b></span></div><div className="topbar-meta"><span className="live-dot"/> LIVE ANALYTICS <span className="divider"/> PLANT 03 <span className="avatar">JK</span></div></header>
 <section className="hero"><div><p className="eyebrow">QUALITY CONTROL / 03</p><h1>품질 데이터 <em>인사이트</em></h1><p className="hero-copy">지종·평량 계층으로 필터링하고 위치별 측정값 평균을 확인하세요.</p></div><div className="hero-stamp"><span>LAST SYNC</span><strong>08.10.2026</strong><small>09:04 KST</small></div></section>
 <section className="upload-zone-wrap"><div className={`upload-zone ${dragging?'is-dragging':''}`} onDragOver={e=>{e.preventDefault();setDragging(true)}} onDragLeave={()=>setDragging(false)} onDrop={e=>{e.preventDefault();setDragging(false);void loadFile(e.dataTransfer.files[0])}} onClick={()=>inputRef.current?.click()}><input ref={inputRef} type="file" accept=".xls,.xlsx" hidden onChange={e=>void loadFile(e.target.files?.[0])}/><span className="upload-icon"><UploadCloud size={22}/></span><div><strong>엑셀 파일을 여기에 놓으세요</strong><span>.XLS 또는 .XLSX · 변경일지 시트 자동 인식</span></div><button className="ghost-button" type="button">파일 선택 <ChevronDown size={15}/></button></div><div className="file-chip"><FileSpreadsheet size={16}/><span>{fileName}</span><b>{rows.length} rows</b><X size={15}/></div></section>
 <section className="control-bar"><div className="control-label"><Filter size={17}/> FILTERS</div><label>지종 대분류 <select value={grade} onChange={e=>setGrade(e.target.value)}>{grades.map(x=><option key={x}>{x}</option>)}</select></label><label>평량 중분류 <select value={basis} onChange={e=>setBasis(e.target.value)}>{bases.map(x=><option key={x}>{x==='전체 평량'?x:`${x} g/㎡`}</option>)}</select></label><label>품질 항목 <select value={metric} onChange={e=>setMetricSafe(e.target.value)}>{metricOptions.map(x=><option key={x.key} value={x.key}>{x.label}</option>)}</select></label><div className="spec-controls"><span>SPEC LIMITS</span><label className="limit-input usl"><i/> USL <input value={usl} onChange={e=>setUsl(e.target.value)}/></label><label className="limit-input lsl"><i/> LSL <input value={lsl} onChange={e=>setLsl(e.target.value)}/></label></div></section>
 <section className="stat-grid"><article className="stat-card"><span>필터 데이터</span><strong>{filtered.length}<small> 건</small></strong><p><BarChart3 size={14}/> 지종 {grade} · 평량 {basis}</p></article><article className="stat-card"><span>평균 {selected.label}</span><strong>{average.toFixed(metric==='수분'?2:1)}<small> {selected.unit}</small></strong><p className="positive"><Gauge size={14}/> 전·중·후 위치 평균</p></article><article className="stat-card alert-card"><span>규격 이탈</span><strong>{outliers.length}<small> 건</small></strong><p><AlertTriangle size={14}/> 즉시 확인 필요</p></article><article className="stat-card"><span>합격률</span><strong>{passRate}<small>%</small></strong><p className="positive"><CheckCircle2 size={14}/> 양호한 공정 상태</p></article></section>
 <section className="content-grid"><article className="panel chart-panel"><div className="panel-head"><div><p className="panel-kicker">TREND MONITOR</p><h2>{selected.label} 위치 평균 트렌드</h2></div><div className="legend"><span><i className="legend-line"/> 위치 평균</span><span><i className="legend-usl"/> USL / LSL</span></div></div><div className="chart-wrap"><ResponsiveContainer width="100%" height="100%"><AreaChart data={chartData} margin={{top:18,right:18,left:-18,bottom:0}}><defs><linearGradient id="qualityFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#27d3a2" stopOpacity={.27}/><stop offset="100%" stopColor="#27d3a2" stopOpacity={0}/></linearGradient></defs><CartesianGrid stroke="#e8edf2" vertical={false}/><XAxis dataKey="label" tick={{fontSize:10,fill:'#84909d'}} tickLine={false} axisLine={false}/><YAxis domain={[Number(lsl)-6,Number(usl)+6]} tick={{fontSize:10,fill:'#84909d'}} tickLine={false} axisLine={false}/><Tooltip formatter={(v)=>[`${Number(v).toFixed(2)} ${selected.unit}`,selected.label]} contentStyle={{border:0,borderRadius:10}}/><ReferenceLine y={Number(usl)} stroke="#f36a5d" strokeDasharray="5 5" label={{value:'USL',fill:'#f36a5d',fontSize:11,position:'insideTopRight'}}/><ReferenceLine y={Number(lsl)} stroke="#f36a5d" strokeDasharray="5 5" label={{value:'LSL',fill:'#f36a5d',fontSize:11,position:'insideBottomRight'}}/><Area type="monotone" dataKey="value" stroke="#16b98d" strokeWidth={2.5} fill="url(#qualityFill)" dot={(p:any)=>p.payload.row.averages[metric]>Number(usl)||p.payload.row.averages[metric]<Number(lsl)?<circle cx={p.cx} cy={p.cy} r={5} fill="#f36a5d" stroke="#fff" strokeWidth={2}/>:<circle cx={p.cx} cy={p.cy} r={3} fill="#16b98d" stroke="#fff" strokeWidth={2}/>} activeDot={{r:6}}/></AreaChart></ResponsiveContainer></div></article><article className="panel insight-panel"><div className="panel-head"><div><p className="panel-kicker">POSITION AVERAGE</p><h2>전·중·후 평균값</h2></div><Layers3 size={20} className="muted-icon"/></div><div className="insight-highlight"><span className="signal-icon"><Zap size={17} fill="currentColor"/></span><div><strong>{selected.label} 평균 {average.toFixed(2)} {selected.unit}</strong><p>전·중·후 측정값을 평균하여 표시합니다.</p></div></div><div className="insight-list">{positionKeys.map(location=><div key={location}><span>{location} 위치</span><strong>{locationAverage(filtered,metric,location)?.toFixed(2)??'—'} {selected.unit}</strong></div>)}</div></article></section>
 <section className="panel table-panel"><div className="panel-head"><div><p className="panel-kicker">RECENT OBSERVATIONS</p><h2>위치별 원시값과 평균</h2></div><span className="row-count">{filtered.length} records</span></div><div className="table-scroll"><table><thead><tr><th>측정 시각</th><th>거래처 / 항차번호</th><th>지종</th><th>평량</th><th>전</th><th>중</th><th>후</th><th>평균</th><th>판정</th></tr></thead><tbody>{filtered.map(row=>{const val=row.averages[metric];const bad=val!==null&&(val>Number(usl)||val<Number(lsl));return <tr key={row.id}><td className="mono">{row.date} <span>{row.time}</span></td><td><strong>{row.client}</strong><small>{row.source}</small></td><td><span className="grade-tag">{row.grade}</span></td><td className="mono">{row.basisWeight??'—'}</td>{positionKeys.map(k=><td key={k} className="value-cell">{row.values[metric]?.[k]??'—'}</td>)}<td className={bad?'danger-text value-cell':'value-cell'}>{val===null?'—':val.toFixed(2)}</td><td><span className={`status ${bad?'status-alert':'status-ok'}`}>{bad?<AlertTriangle size={13}/>:<CheckCircle2 size={13}/>} {bad?'규격 이탈':'합격'}</span></td></tr>})}</tbody></table></div></section><footer><span>QUALITY LEDGER / PLANT 03</span><span>Data integrity monitored · v1.1.0</span></footer></main>
}
createRoot(document.getElementById('root')!).render(<App/>)
