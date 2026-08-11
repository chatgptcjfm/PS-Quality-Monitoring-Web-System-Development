/** @jsxImportSource react */
import React, { useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import {
  Area, AreaChart, CartesianGrid, ResponsiveContainer, ReferenceLine,
  Tooltip, XAxis, YAxis,
} from 'recharts'
import {
  AlertTriangle, BarChart3, CheckCircle2, ChevronDown, FileSpreadsheet,
  Filter, Gauge, Layers3, UploadCloud, X, Zap,
} from 'lucide-react'

type QualityRow = {
  id: number; date: string; client: string; grade: string; measure: string
  value: number; status: '정상' | '이탈'; time: string; source: string
}

const demoRows: QualityRow[] = [
  { id: 1, date: '08.10', client: '중국 · 제약케이스', grade: 'IV', measure: '평량', value: 303, status: '정상', time: '07:38', source: '6810J30001' },
  { id: 2, date: '08.10', client: '중국 · 제약케이스', grade: 'IV', measure: '평량', value: 304, status: '정상', time: '07:44', source: '6810J30003' },
  { id: 3, date: '08.10', client: '신승아이엔씨', grade: 'IV', measure: '평량', value: 305, status: '정상', time: '07:49', source: '6810J30005' },
  { id: 4, date: '08.10', client: '디와이 · 흥아', grade: 'IV', measure: '평량', value: 302, status: '정상', time: '08:01', source: '6810J30009' },
  { id: 5, date: '08.10', client: '중국 · 제약케이스', grade: 'IV', measure: '평량', value: 307, status: '이탈', time: '08:16', source: '6810J30012' },
  { id: 6, date: '08.10', client: '태국 · 포장용지', grade: 'III', measure: '평량', value: 298, status: '정상', time: '08:32', source: '6810J30018' },
  { id: 7, date: '08.10', client: '중국 · 제약케이스', grade: 'IV', measure: '평량', value: 301, status: '정상', time: '08:48', source: '6810J30020' },
  { id: 8, date: '08.10', client: '신승아이엔씨', grade: 'IV', measure: '평량', value: 296, status: '이탈', time: '09:04', source: '6810J30022' },
]

const metricOptions = [
  { label: '평량 (g/㎡)', key: '평량', usl: 306, lsl: 294, unit: 'g/㎡' },
  { label: '두께 (㎛)', key: '두께', usl: 360, lsl: 340, unit: '㎛' },
  { label: '수분 (%)', key: '수분', usl: 7.5, lsl: 4.5, unit: '%' },
]

function parseWorkbook(file: File): Promise<QualityRow[]> {
  return file.arrayBuffer().then((buffer) => {
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true })
    const sheetName = workbook.SheetNames.find((name) => name.includes('변경일지')) ?? workbook.SheetNames[0]
    const rows = XLSX.utils.sheet_to_json<(string | number | Date | null)[]>(workbook.Sheets[sheetName], { header: 1, defval: null })
    const output: QualityRow[] = []
    rows.slice(8).forEach((row, index) => {
      const client = String(row[5] ?? '').replace(/\n/g, ' ').trim()
      const grade = String(row[6] ?? '').trim()
      const value = Number(row[8])
      if (!client && !grade && !Number.isFinite(value)) return
      const date = `${String(row[1] ?? '').padStart(2, '0')}.${String(row[2] ?? '').padStart(2, '0')}`
      const time = `${String(row[3] ?? '').padStart(2, '0')}:${String(row[4] ?? '').padStart(2, '0')}`
      output.push({
        id: index + 1, date, client: client || '미지정 거래처', grade: grade || '미지정', measure: '평량', value,
        status: value > 306 || value < 294 ? '이탈' : '정상', time, source: client.slice(0, 12) || `ROW-${index + 1}`,
      })
    })
    return output
  })
}

function App() {
  const inputRef = useRef<HTMLInputElement>(null)
  const [rows, setRows] = useState<QualityRow[]>(demoRows)
  const [fileName, setFileName] = useState('3호기 2026.08.10.xls')
  const [grade, setGrade] = useState('전체 지종')
  const [metric, setMetric] = useState(metricOptions[0].key)
  const [usl, setUsl] = useState('306')
  const [lsl, setLsl] = useState('294')
  const [dragging, setDragging] = useState(false)
  const selectedMetric = metricOptions.find((item) => item.key === metric) ?? metricOptions[0]
  const grades = ['전체 지종', ...Array.from(new Set(rows.map((row) => row.grade)))]

  const filteredRows = useMemo(() => rows.filter((row) => grade === '전체 지종' || row.grade === grade), [rows, grade])
  const chartData = filteredRows.map((row) => ({ ...row, label: `${row.date} ${row.time}`, value: Number(row.value) }))
  const outliers = filteredRows.filter((row) => row.value > Number(usl) || row.value < Number(lsl))
  const avg = filteredRows.length ? filteredRows.reduce((sum, row) => sum + row.value, 0) / filteredRows.length : 0
  const passRate = filteredRows.length ? Math.round(((filteredRows.length - outliers.length) / filteredRows.length) * 100) : 0

  const loadFile = async (file?: File) => {
    if (!file || !/\.xls[x]?$/.test(file.name.toLowerCase())) return
    const parsed = await parseWorkbook(file)
    setRows(parsed.length ? parsed : demoRows)
    setFileName(file.name)
  }

  const handleDrop = (event: React.DragEvent) => {
    event.preventDefault(); setDragging(false); void loadFile(event.dataTransfer.files[0])
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand"><span className="brand-mark"><Zap size={16} fill="currentColor" /></span><span>QUALITY <b>LEDGER</b></span></div>
        <div className="topbar-meta"><span className="live-dot" /> LIVE ANALYTICS <span className="divider" /> PLANT 03 <span className="avatar">JK</span></div>
      </header>

      <section className="hero">
        <div><p className="eyebrow">QUALITY CONTROL / 03</p><h1>품질 데이터 <em>인사이트</em></h1><p className="hero-copy">변경일지 데이터에서 패턴을 읽고, 규격 이탈을 한눈에 관리하세요.</p></div>
        <div className="hero-stamp"><span>LAST SYNC</span><strong>08.10.2026</strong><small>09:04 KST</small></div>
      </section>

      <section className="upload-zone-wrap">
        <div className={`upload-zone ${dragging ? 'is-dragging' : ''}`} onDragOver={(e) => { e.preventDefault(); setDragging(true) }} onDragLeave={() => setDragging(false)} onDrop={handleDrop} onClick={() => inputRef.current?.click()}>
          <input ref={inputRef} type="file" accept=".xls,.xlsx" hidden onChange={(e) => void loadFile(e.target.files?.[0])} />
          <span className="upload-icon"><UploadCloud size={22} /></span>
          <div><strong>엑셀 파일을 여기에 놓으세요</strong><span>.XLS 또는 .XLSX · 변경일지 시트 자동 인식</span></div>
          <button className="ghost-button" type="button">파일 선택 <ChevronDown size={15} /></button>
        </div>
        <div className="file-chip"><FileSpreadsheet size={16} /><span>{fileName}</span><b>{rows.length} rows</b><X size={15} /></div>
      </section>

      <section className="control-bar">
        <div className="control-label"><Filter size={17} /> FILTERS</div>
        <label>지종 <select value={grade} onChange={(e) => setGrade(e.target.value)}>{grades.map((item) => <option key={item}>{item}</option>)}</select></label>
        <label>측정 항목 <select value={metric} onChange={(e) => setMetric(e.target.value)}>{metricOptions.map((item) => <option key={item.key} value={item.key}>{item.label}</option>)}</select></label>
        <div className="spec-controls"><span>SPEC LIMITS</span><label className="limit-input usl"><i /> USL <input value={usl} onChange={(e) => setUsl(e.target.value)} /></label><label className="limit-input lsl"><i /> LSL <input value={lsl} onChange={(e) => setLsl(e.target.value)} /></label></div>
      </section>

      <section className="stat-grid">
        <article className="stat-card"><span>측정 데이터</span><strong>{filteredRows.length.toLocaleString()} <small>건</small></strong><p><BarChart3 size={14} /> 최근 변경일지 기준</p></article>
        <article className="stat-card"><span>평균 {metric}</span><strong>{avg.toFixed(metric === '수분' ? 2 : 1)} <small>{selectedMetric.unit}</small></strong><p className="positive"><Gauge size={14} /> 공정 중심값</p></article>
        <article className="stat-card alert-card"><span>규격 이탈</span><strong>{outliers.length} <small>건</small></strong><p><AlertTriangle size={14} /> 즉시 확인 필요</p></article>
        <article className="stat-card"><span>합격률</span><strong>{passRate}<small>%</small></strong><p className="positive"><CheckCircle2 size={14} /> 양호한 공정 상태</p></article>
      </section>

      <section className="content-grid">
        <article className="panel chart-panel"><div className="panel-head"><div><p className="panel-kicker">TREND MONITOR</p><h2>시계열 품질 트렌드</h2></div><div className="legend"><span><i className="legend-line" /> 측정값</span><span><i className="legend-usl" /> USL / LSL</span></div></div><div className="chart-wrap"><ResponsiveContainer width="100%" height="100%"><AreaChart data={chartData} margin={{ top: 18, right: 18, left: -18, bottom: 0 }}><defs><linearGradient id="qualityFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#27d3a2" stopOpacity={0.27} /><stop offset="100%" stopColor="#27d3a2" stopOpacity={0} /></linearGradient></defs><CartesianGrid stroke="#e8edf2" vertical={false} /><XAxis dataKey="label" tick={{ fontSize: 10, fill: '#84909d' }} tickLine={false} axisLine={false} /><YAxis domain={[Number(lsl) - 6, Number(usl) + 6]} tick={{ fontSize: 10, fill: '#84909d' }} tickLine={false} axisLine={false} /><Tooltip contentStyle={{ border: '0', borderRadius: 10, boxShadow: '0 8px 30px rgba(20,40,70,.12)' }} formatter={(value) => [`${value} ${selectedMetric.unit}`, metric]} /><ReferenceLine y={Number(usl)} stroke="#f36a5d" strokeDasharray="5 5" label={{ value: 'USL', fill: '#f36a5d', fontSize: 11, position: 'insideTopRight' }} /><ReferenceLine y={Number(lsl)} stroke="#f36a5d" strokeDasharray="5 5" label={{ value: 'LSL', fill: '#f36a5d', fontSize: 11, position: 'insideBottomRight' }} /><Area type="monotone" dataKey="value" stroke="#16b98d" strokeWidth={2.5} fill="url(#qualityFill)" dot={(props) => props.payload.status === '이탈' ? <circle {...props} r={5} fill="#f36a5d" stroke="#fff" strokeWidth={2} /> : <circle {...props} r={3} fill="#16b98d" stroke="#fff" strokeWidth={2} />} activeDot={{ r: 6 }} /></AreaChart></ResponsiveContainer></div></article>

        <article className="panel insight-panel"><div className="panel-head"><div><p className="panel-kicker">AI QUALITY READOUT</p><h2>공정 해석</h2></div><Layers3 size={20} className="muted-icon" /></div><div className="insight-highlight"><span className="signal-icon"><Zap size={17} fill="currentColor" /></span><div><strong>안정적인 공정 흐름</strong><p>최근 측정값이 중심선을 유지하고 있습니다.</p></div></div><div className="insight-list"><div><span>중심 경향</span><strong>{avg.toFixed(1)} {selectedMetric.unit}</strong></div><div><span>최대 편차</span><strong className={outliers.length ? 'danger-text' : ''}>{outliers.length ? `${Math.max(...filteredRows.map((r) => Math.abs(r.value - avg))).toFixed(1)} ${selectedMetric.unit}` : '없음'}</strong></div><div><span>샘플 범위</span><strong>{filteredRows.length ? `${Math.min(...filteredRows.map(r => r.value))} — ${Math.max(...filteredRows.map(r => r.value))}` : '—'}</strong></div></div><button className="text-button" type="button">상세 리포트 보기 <span>→</span></button></article>
      </section>

      <section className="panel table-panel"><div className="panel-head"><div><p className="panel-kicker">RECENT OBSERVATIONS</p><h2>최근 측정 데이터</h2></div><span className="row-count">{filteredRows.length} records</span></div><div className="table-scroll"><table><thead><tr><th>측정 시각</th><th>거래처 / 항차번호</th><th>지종</th><th>항목</th><th>측정값</th><th>판정</th></tr></thead><tbody>{filteredRows.map((row) => { const isOutlier = row.value > Number(usl) || row.value < Number(lsl); return <tr key={row.id}><td className="mono">{row.date} <span>{row.time}</span></td><td><strong>{row.client}</strong><small>{row.source}</small></td><td><span className="grade-tag">{row.grade}</span></td><td>{metric}</td><td className={isOutlier ? 'danger-text value-cell' : 'value-cell'}>{row.value} <small>{selectedMetric.unit}</small></td><td><span className={`status ${isOutlier ? 'status-alert' : 'status-ok'}`}>{isOutlier ? <AlertTriangle size={13} /> : <CheckCircle2 size={13} />}{isOutlier ? '규격 이탈' : '합격'}</span></td></tr>})}</tbody></table></div></section>
      <footer><span>QUALITY LEDGER / PLANT 03</span><span>Data integrity monitored · v1.0.0</span></footer>
    </main>
  )
}

createRoot(document.getElementById('root')!).render(<App />)
