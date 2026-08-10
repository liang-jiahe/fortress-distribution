import { useEffect, useMemo, useRef, useState } from 'react'
import ExcelJS from 'exceljs'

type Member = {
  id: string
  name: string
  power: number
  score: number | null
  remark: string
  order: number
}

type PackageType = 'fire' | 'mid1' | 'mid2'
type Schedule = Record<string, string | null>

type Session = {
  id: string
  label: string
  block: 'pink' | 'orange' | 'yellow' | 'green' | 'cyan' | 'blue'
  startColumn: number
  startRow: number
}

const PACKAGE_LABELS: Record<PackageType, string> = { fire: '火', mid1: '中一', mid2: '中二' }
const TYPE_ORDER: PackageType[] = ['fire', 'mid1', 'mid2']
const TIERS = [
  { min: 1, max: 5, fire: 2, middle: 3, coins: 62, color: 'pink' },
  { min: 6, max: 10, fire: 2, middle: 2, coins: 60, color: 'orange' },
  { min: 11, max: 15, fire: 2, middle: 1, coins: 58, color: 'yellow' },
  { min: 16, max: 20, fire: 1, middle: 3, coins: 55, color: 'green' },
  { min: 21, max: 25, fire: 1, middle: 2, coins: 53, color: 'cyan' },
  { min: 26, max: 30, fire: 0, middle: 5, coins: 52, color: 'blue' },
] as const

const SESSIONS: Session[] = [
  { id: 'sat-pm', label: '周六下', block: 'pink', startColumn: 1, startRow: 2 },
  { id: 'sun', label: '周日', block: 'pink', startColumn: 4, startRow: 2 },
  { id: 'mon', label: '周一', block: 'orange', startColumn: 7, startRow: 2 },
  { id: 'tue', label: '周二', block: 'green', startColumn: 10, startRow: 2 },
  { id: 'wed', label: '周三', block: 'yellow', startColumn: 1, startRow: 9 },
  { id: 'thu', label: '周四', block: 'yellow', startColumn: 4, startRow: 9 },
  { id: 'fri', label: '周五', block: 'cyan', startColumn: 7, startRow: 9 },
  { id: 'sat-am', label: '星期六上', block: 'green', startColumn: 10, startRow: 9 },
]

const SAMPLE_NAMES = ['太初星', '关注塔菲喵', '花云青', '无双', '御茨星', '念君夏', '夏弥', '谦灵星', '小星星', '沈七涵', '晓行星', '心芯星', '猫猫星', '绝地', '鸿鹄', '云岫', '亿丈龙我', '椛七', '超级萝卜大王', '拳王', '浅帐星', '别急稳一手', '小苏在这里', '时愿星', '伦敦街尾吻别', '33', '季时雨花知否', '我一直都在', '弦', '白慕']
const SAMPLE_POWER = [3852, 3751, 3736, 3258, 3220, 3077, 2867, 2881, 2802, 2746, 2732, 2684, 2633, 2567, 2547, 2542, 2469, 2460, 2408, 2405, 2323, 2314, 2306, 2294, 2218, 2206, 2193, 2189, 2110, 2096]
const SAMPLE_SCORE_BY_NAME: Record<string, number> = {
  '太初星': 36, '关注塔菲喵': 36, '花云青': 36, '无双': 36, '御茨星': 35, '念君夏': 36, '夏弥': 36, '谦灵星': 36, '小星星': 35, '沈七涵': 16, '晓行星': 36, '心芯星': 36, '猫猫星': 36, '绝地': 36, '鸿鹄': 26, '云岫': 26, '亿丈龙我': 36, '椛七': 36, '超级萝卜大王': 5, '拳王': 36, '浅帐星': 36, '别急稳一手': 36, '小苏在这里': 15, '时愿星': 17, '伦敦街尾吻别': 36, '33': 36, '季时雨花知否': 36, '我一直都在': 15, '弦': 26, '白慕': 36,
}

function makeSampleMembers(): Member[] {
  return SAMPLE_NAMES.map((name, index) => ({ id: `m-${index + 1}`, name, power: SAMPLE_POWER[index], score: SAMPLE_SCORE_BY_NAME[name] ?? null, remark: '', order: index }))
}

function tierForRank(rank: number) { return TIERS.find((tier) => rank >= tier.min && rank <= tier.max) ?? TIERS[TIERS.length - 1] }
function rankMembers(members: Member[]) { return [...members].sort((a, b) => (b.score ?? -1) - (a.score ?? -1) || b.power - a.power || a.order - b.order) }
function powerRankMembers(members: Member[]) { return [...members].sort((a, b) => b.power - a.power || a.order - b.order) }
function cellKey(sessionId: string, type: PackageType, row: number) { return `${sessionId}:${type}:${row}` }
function setGroup(schedule: Schedule, sessionId: string, type: PackageType, ids: (string | null)[]) { ids.forEach((id, row) => { schedule[cellKey(sessionId, type, row)] = id }) }

function buildAutoSchedule(ranked: Member[]): Schedule {
  const schedule: Schedule = {}
  SESSIONS.forEach((session) => TYPE_ORDER.forEach((type) => setGroup(schedule, session.id, type, [null, null, null, null, null])))
  const group = (from: number) => ranked.slice(from, from + 5).map((member) => member?.id ?? null)
  setGroup(schedule, 'sat-pm', 'fire', group(0)); setGroup(schedule, 'sat-pm', 'mid1', group(0)); setGroup(schedule, 'sat-pm', 'mid2', group(0))
  setGroup(schedule, 'sun', 'fire', group(0)); setGroup(schedule, 'sun', 'mid1', group(0)); setGroup(schedule, 'sun', 'mid2', group(10))
  setGroup(schedule, 'mon', 'fire', group(5)); setGroup(schedule, 'mon', 'mid1', group(5)); setGroup(schedule, 'mon', 'mid2', group(5))
  setGroup(schedule, 'tue', 'fire', group(5)); setGroup(schedule, 'tue', 'mid1', group(15)); setGroup(schedule, 'tue', 'mid2', group(25))
  setGroup(schedule, 'wed', 'fire', group(10)); setGroup(schedule, 'wed', 'mid1', group(25)); setGroup(schedule, 'wed', 'mid2', group(25))
  setGroup(schedule, 'thu', 'fire', group(10)); setGroup(schedule, 'thu', 'mid1', group(25)); setGroup(schedule, 'thu', 'mid2', group(25))
  setGroup(schedule, 'fri', 'fire', group(20)); setGroup(schedule, 'fri', 'mid1', group(20)); setGroup(schedule, 'fri', 'mid2', group(20))
  setGroup(schedule, 'sat-am', 'fire', group(15)); setGroup(schedule, 'sat-am', 'mid1', group(15)); setGroup(schedule, 'sat-am', 'mid2', group(15))
  return schedule
}

function parseCell(value: unknown): string { return value == null ? '' : String(value).trim() }
function numeric(value: unknown): number | null { const n = Number(value); return Number.isFinite(n) ? n : null }

async function importWorkbook(file: File): Promise<Member[]> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await file.arrayBuffer())
  const byName = new Map<string, Partial<Member> & { order: number }>()
  let order = 0
  const absorb = (name: unknown, power: unknown, score: unknown, remark: unknown) => {
    const clean = parseCell(name)
    if (!clean || clean === '合计' || clean === '人员') return
    const existing = byName.get(clean) ?? { order: order++ }
    const p = numeric(power); const s = numeric(score)
    if (p != null) existing.power = p
    if (s != null) existing.score = s
    if (parseCell(remark)) existing.remark = parseCell(remark)
    byName.set(clean, existing)
  }
  const readRankSheet = (sheet: ExcelJS.Worksheet, mode: 'power' | 'score') => {
    let header = -1; let nameCol = -1; let valueCol = -1; let remarkCol = -1
    sheet.eachRow((row, rowNumber) => {
      const values = row.values as unknown[]
      values.forEach((value, col) => { if (parseCell(value) === '人员') { header = rowNumber; nameCol = col } if (parseCell(value) === (mode === 'power' ? '战力' : '分数')) valueCol = col; if (parseCell(value) === '备注') remarkCol = col })
    })
    if (header < 0 || nameCol < 0) return
    for (let r = header + 1; r <= sheet.rowCount; r += 1) { const row = sheet.getRow(r); absorb(row.getCell(nameCol).value, mode === 'power' ? row.getCell(valueCol).value : undefined, mode === 'score' ? row.getCell(valueCol).value : undefined, remarkCol > 0 ? row.getCell(remarkCol).value : undefined) }
  }
  const powerSheet = workbook.getWorksheet('战力排名')
  const scoreSheet = workbook.getWorksheet('分数排名')
  if (powerSheet) readRankSheet(powerSheet, 'power')
  if (scoreSheet) readRankSheet(scoreSheet, 'score')
  const first = workbook.worksheets[0]
  if (first) {
    for (let r = 1; r <= first.rowCount; r += 1) {
      const row = first.getRow(r)
      for (let c = 1; c <= row.cellCount; c += 1) {
        if (parseCell(row.getCell(c).value) === '人员' && parseCell(row.getCell(c + 1).value) === '分数') {
          for (let rr = r + 1; rr <= Math.min(r + 31, first.rowCount); rr += 1) absorb(first.getRow(rr).getCell(c).value, undefined, first.getRow(rr).getCell(c + 1).value, first.getRow(rr).getCell(c + 2).value)
        }
      }
    }
  }
  return [...byName.entries()].map(([name, data], index) => ({ id: `m-${Date.now()}-${index}`, name, power: data.power ?? 0, score: data.score ?? null, remark: data.remark ?? '', order: data.order ?? index })).sort((a, b) => a.order - b.order)
}

function colorHex(color: string) { return ({ pink: 'F7DDE3', orange: 'FFC000', yellow: 'FFF200', green: '92D050', cyan: '10B8E8', blue: '4472C4' } as Record<string, string>)[color] ?? 'FFFFFF' }

async function exportWorkbook(members: Member[], schedule: Schedule, ranked: Member[]) {
  const workbook = new ExcelJS.Workbook(); const sheet = workbook.addWorksheet('繁星本周要塞包分配')
  sheet.properties.defaultRowHeight = 20
  for (let col = 1; col <= 12; col += 1) sheet.getColumn(col).width = 14
  sheet.getColumn(13).width = 8; sheet.getColumn(14).width = 22; sheet.getColumn(15).width = 10; sheet.getColumn(16).width = 20
  const border = { top: { style: 'thin' as const, color: { argb: 'FF3F3F3F' } }, left: { style: 'thin' as const, color: { argb: 'FF3F3F3F' } }, bottom: { style: 'thin' as const, color: { argb: 'FF3F3F3F' } }, right: { style: 'thin' as const, color: { argb: 'FF3F3F3F' } } }
  const center = { vertical: 'middle' as const, horizontal: 'center' as const }
  sheet.mergeCells('A1:P1'); sheet.getCell('A1').value = '繁星本周要塞包分配'; sheet.getCell('A1').font = { name: '宋体', size: 16, bold: true }; sheet.getCell('A1').alignment = center; sheet.getRow(1).height = 34
  const rankHeader = ['本周', '人员', '分数', '备注']; ['M2', 'N2', 'O2', 'P2'].forEach((cell, i) => { sheet.getCell(cell).value = rankHeader[i]; sheet.getCell(cell).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } }; sheet.getCell(cell).font = { name: '宋体', size: 12 }; sheet.getCell(cell).alignment = center; sheet.getCell(cell).border = border })
  const memberById = new Map(members.map((member) => [member.id, member]))
  SESSIONS.forEach((session) => {
    const start = session.startColumn; const end = start + 2; const headerRow = session.startRow; const dataStart = headerRow + 2
    sheet.mergeCells(headerRow, start, headerRow, end); sheet.getCell(headerRow, start).value = session.label; sheet.getCell(headerRow, start).alignment = center; sheet.getCell(headerRow, start).font = { name: '楷体', size: 12 }; sheet.getCell(headerRow, start).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } }
    TYPE_ORDER.forEach((type, offset) => { const assigned = memberById.get(schedule[cellKey(session.id, type, 0)] ?? ''); const assignedRank = assigned ? ranked.findIndex((entry) => entry.id === assigned.id) + 1 : 0; const cell = sheet.getCell(headerRow + 1, start + offset); cell.value = PACKAGE_LABELS[type]; cell.alignment = center; cell.border = border; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colorHex(assignedRank ? tierForRank(assignedRank).color : session.block)}` } } })
    for (let row = 0; row < 5; row += 1) {
      const values = TYPE_ORDER.map((type) => memberById.get(schedule[cellKey(session.id, type, row)] ?? '')?.name ?? '')
      TYPE_ORDER.forEach((type, offset) => { const cell = sheet.getCell(dataStart + row, start + offset); cell.value = values[offset]; cell.alignment = center; cell.border = border; const member = memberById.get(schedule[cellKey(session.id, type, row)] ?? ''); const rank = member ? ranked.findIndex((entry) => entry.id === member.id) + 1 : 0; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${rank ? colorHex(tierForRank(rank).color) : 'FFFFFF'}` } } })
      if (values[0] && values[0] === values[1] && values[1] === values[2]) { sheet.mergeCells(dataStart + row, start, dataStart + row, end); sheet.getCell(dataStart + row, start).value = values[0] } else if (values[0] && values[0] === values[1]) { sheet.mergeCells(dataStart + row, start, dataStart + row, start + 1); sheet.getCell(dataStart + row, start).value = values[0] }
    }
  })
  sheet.mergeCells('A17:L17'); sheet.getCell('A17').value = '分配方案：每周1火2中，每人每周可得42个币，另外火包7币，中包2币，按每周考核分来分包'; sheet.getCell('A17').alignment = center; sheet.getCell('A17').font = { name: '宋体', size: 11 }
  TIERS.forEach((tier, index) => { const row = 18 + index; sheet.mergeCells(row, 1, row, 3); sheet.mergeCells(row, 4, row, 9); sheet.mergeCells(row, 10, row, 12); sheet.getCell(row, 1).value = `考核${tier.min}-${tier.max}`; sheet.getCell(row, 4).value = `每周${tier.fire}火、${tier.middle}中`; sheet.getCell(row, 10).value = `每周${tier.coins}币`; [1, 4, 10].forEach((col) => { sheet.getCell(row, col).alignment = center; sheet.getCell(row, col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colorHex(tier.color)}` } }; sheet.getCell(row, col).border = border }) })
  sheet.mergeCells('A24:L24'); sheet.getCell('A24').value = '特殊情况：扣包/或特殊情况奖励包'; sheet.getCell('A24').alignment = center; sheet.getCell('A24').border = border
  ranked.slice(0, 30).forEach((member, index) => { const row = index + 3; const rankCell = sheet.getCell(row, 13); rankCell.value = index + 1; rankCell.alignment = center; rankCell.border = border; rankCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colorHex(tierForRank(index + 1).color)}` } }; [member.name, member.score ?? '', member.remark].forEach((value, i) => { const cell = sheet.getCell(row, 14 + i); cell.value = value; cell.alignment = center; cell.border = border; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colorHex(tierForRank(index + 1).color)}` } } }) })
  sheet.mergeCells('A26:L26'); sheet.getCell('A26').value = '战力、考核分排名每周六晚统计，每周更新'; sheet.getCell('A26').alignment = center
  const buffer = await workbook.xlsx.writeBuffer(); const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); const url = URL.createObjectURL(blob); const link = document.createElement('a'); link.href = url; link.download = '繁星本周要塞包分配.xlsx'; link.click(); URL.revokeObjectURL(url)
}

export default function App() {
  const [members, setMembers] = useState<Member[]>(() => { try { const saved = localStorage.getItem('fortress-members'); return saved ? JSON.parse(saved) : makeSampleMembers() } catch { return makeSampleMembers() } })
  const [contest, setContest] = useState(false); const [search, setSearch] = useState(''); const [manualSchedule, setManualSchedule] = useState<Schedule | null>(null); const [notice, setNotice] = useState('已加载示例数据，可直接编辑或导入本周表格。'); const [activeSection, setActiveSection] = useState('matrix'); const fileRef = useRef<HTMLInputElement>(null)
  const ranked = useMemo(() => rankMembers(members), [members]); const powerRanked = useMemo(() => powerRankMembers(members), [members]); const autoSchedule = useMemo(() => buildAutoSchedule(ranked), [ranked]); const schedule = manualSchedule ?? autoSchedule; const scoreMax = contest ? 57 : 37
  const counts = useMemo(() => { const result = new Map<string, { fire: number; middle: number }>(); Object.entries(schedule).forEach(([key, id]) => { if (!id) return; const type = key.split(':')[1] as PackageType; const current = result.get(id) ?? { fire: 0, middle: 0 }; if (type === 'fire') current.fire += 1; else current.middle += 1; result.set(id, current) }); return result }, [schedule])
  useEffect(() => { localStorage.setItem('fortress-members', JSON.stringify(members)) }, [members])
  const updateMember = (id: string, patch: Partial<Member>) => { setMembers((current) => current.map((member) => member.id === id ? { ...member, ...patch } : member)); setManualSchedule(null) }
  const addMember = () => { const index = members.length + 1; setMembers((current) => [...current, { id: `m-${Date.now()}`, name: `新成员${index}`, power: 0, score: null, remark: '', order: current.length }]); setManualSchedule(null); setNotice('已新增成员，请填写姓名、战力和考核分。') }
  const removeMember = (id: string) => { setMembers((current) => current.filter((member) => member.id !== id)); setManualSchedule(null) }
  const changeCell = (sessionId: string, type: PackageType, row: number, id: string) => { const next = { ...(manualSchedule ?? autoSchedule) }; next[cellKey(sessionId, type, row)] = id || null; setManualSchedule(next) }
  const visibleMembers = members.filter((member) => member.name.toLowerCase().includes(search.toLowerCase()))
  const reset = () => { setMembers(makeSampleMembers()); setManualSchedule(null); setNotice('已恢复示例数据。') }
  const clear = () => { setMembers([]); setManualSchedule(null); setNotice('已清空成员数据。') }
  const handleImport = async (event: React.ChangeEvent<HTMLInputElement>) => { const file = event.target.files?.[0]; if (!file) return; try { const imported = await importWorkbook(file); if (!imported.length) throw new Error('没有识别到成员'); setMembers(imported); setManualSchedule(null); setNotice(`已导入 ${imported.length} 名成员。`) } catch (error) { setNotice(`导入失败：${error instanceof Error ? error.message : '文件格式不正确'}`) } finally { event.target.value = '' } }
  const goTo = (section: string) => { setActiveSection(section); document.getElementById(section)?.scrollIntoView({ behavior: 'smooth', block: 'start' }) }
  return <div className="app-shell">
    <aside className="sidebar"><div className="brand-mark">🐾</div><h1>繁星要塞</h1><nav><button className={activeSection === 'matrix' ? 'active' : ''} onClick={() => goTo('matrix')}>要塞分包</button><button className={activeSection === 'members' ? 'active' : ''} onClick={() => goTo('members')}>成员管理</button><button className={activeSection === 'ranking' ? 'active' : ''} onClick={() => goTo('members')}>排名与考核</button><button className={activeSection === 'instructions' ? 'active' : ''} onClick={() => goTo('instructions')}>使用说明</button></nav><div className="sidebar-note">数据只保存在当前浏览器<br />导入即用，导出即走</div></aside>
    <main className="content">
      <header className="topbar"><div><div className="eyebrow">FORTRESS DISTRIBUTION</div><h2>本周要塞包分配</h2></div><div className="toolbar"><button className="btn secondary" onClick={() => fileRef.current?.click()}>⇧ 导入 XLSX</button><input ref={fileRef} type="file" accept=".xlsx,.xls" hidden onChange={handleImport} /><button className="btn primary" onClick={() => exportWorkbook(members, schedule, ranked)}>⇩ 导出单表 XLSX</button></div></header>
      <section className="notice">{notice}<span className="notice-right"><label className="toggle"><input type="checkbox" checked={contest} onChange={(e) => setContest(e.target.checked)} /><span></span> 争霸周（最高 {scoreMax} 分）</label></span></section>
      <section className="stats"><div className="stat-card"><span>成员人数</span><strong>{members.length}</strong><small>{members.length === 30 ? '模板完整' : '目标 30 人'}</small></div><div className="stat-card"><span>当前最高分</span><strong>{ranked[0]?.score ?? '—'}</strong><small>普通周 37 · 争霸周 57</small></div><div className="stat-card"><span>火包总量</span><strong>40</strong><small>8 个时段 × 5 份</small></div><div className="stat-card"><span>中包总量</span><strong>80</strong><small>8 个时段 × 10 份</small></div></section>
      <section id="matrix" className="panel matrix-panel"><div className="panel-heading"><div><span className="eyebrow">AUTO LAYOUT</span><h3>彩色分包矩阵</h3></div><button className="btn ghost" onClick={() => { setManualSchedule(null); setNotice('已按当前排名重新生成矩阵。') }}>↻ 重新生成矩阵</button></div><div className="matrix-scroll"><div className="matrix-grid">{SESSIONS.map((session) => <SessionBlock key={session.id} session={session} schedule={schedule} ranked={ranked} members={members} onChange={changeCell} />)}</div></div><div className="legend">{TIERS.map((tier) => <span key={tier.min}><i className={`swatch ${tier.color}`}></i>{tier.min}-{tier.max} 名</span>)}</div></section>
      <section id="members" className="panel"><span id="ranking" className="anchor-target"></span><div className="panel-heading"><div><span className="eyebrow">MEMBER ROSTER</span><h3>成员与考核分</h3></div><div className="row-actions"><input className="search" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="搜索姓名" /><button className="btn primary" onClick={addMember}>＋ 新增成员</button><button className="btn ghost" onClick={reset}>恢复示例</button><button className="btn danger" onClick={clear}>清空</button></div></div><div className="table-wrap"><table><thead><tr><th>战力排名</th><th>成员姓名</th><th>战力</th><th>分包排名</th><th>考核分</th><th>火/中包</th><th>币数</th><th>备注</th><th></th></tr></thead><tbody>{visibleMembers.map((member) => { const powerRank = powerRanked.findIndex((entry) => entry.id === member.id) + 1; const packageRank = ranked.findIndex((entry) => entry.id === member.id) + 1; const tier = packageRank ? tierForRank(packageRank) : TIERS[0]; const count = counts.get(member.id) ?? { fire: 0, middle: 0 }; return <tr key={member.id}><td><span className="rank-pill">{powerRank || '—'}</span></td><td><input value={member.name} onChange={(e) => updateMember(member.id, { name: e.target.value })} /></td><td><input type="number" value={member.power} onChange={(e) => updateMember(member.id, { power: Number(e.target.value) || 0 })} /></td><td><span className={`tier-dot ${tier.color}`}>{packageRank || '—'}</span></td><td><input type="number" min="0" max={scoreMax} value={member.score ?? ''} placeholder="未填" onChange={(e) => updateMember(member.id, { score: e.target.value === '' ? null : Number(e.target.value) })} /></td><td>{count.fire} 火 / {count.middle} 中</td><td>{packageRank ? tier.coins : '—'}</td><td><input value={member.remark} onChange={(e) => updateMember(member.id, { remark: e.target.value })} placeholder="备注" /></td><td><button className="icon-btn" title="删除成员" onClick={() => removeMember(member.id)}>×</button></td></tr> })}</tbody></table>{visibleMembers.length === 0 && <div className="empty">没有匹配的成员，点击“新增成员”开始。</div>}</div></section>
      <section id="instructions" className="panel rules-panel"><div><span className="eyebrow">PACKAGE RULES</span><h3>分包档位说明</h3></div><div className="tier-cards">{TIERS.map((tier) => <div className={`tier-card ${tier.color}`} key={tier.min}><b>{tier.min}-{tier.max}</b><span>{tier.fire} 火 · {tier.middle} 中</span><strong>{tier.coins} 币</strong></div>)}</div><p className="rule-copy">普通周考核满分 37 分；有争霸赛的周满分 57 分。考核分只用于排序，当前没有自动扣包。</p></section>
    </main>
  </div>
}

function SessionBlock({ session, schedule, ranked, members, onChange }: { session: Session; schedule: Schedule; ranked: Member[]; members: Member[]; onChange: (sessionId: string, type: PackageType, row: number, id: string) => void }) {
  const memberById = new Map(members.map((member) => [member.id, member]))
  return <div className={`session-block ${session.block}`}><div className="session-title">{session.label}</div><div className="session-types">{TYPE_ORDER.map((type) => { const id = schedule[cellKey(session.id, type, 0)] ?? ''; const member = memberById.get(id); const rank = member ? ranked.findIndex((entry) => entry.id === member.id) + 1 : 0; return <div className={rank ? `header-fill ${tierForRank(rank).color}` : ''} key={type}>{PACKAGE_LABELS[type]}</div> })}</div>{Array.from({ length: 5 }, (_, row) => <div className="session-row" key={row}>{TYPE_ORDER.map((type) => { const id = schedule[cellKey(session.id, type, row)] ?? ''; const member = memberById.get(id); const rank = member ? ranked.findIndex((entry) => entry.id === member.id) + 1 : 0; return <select key={type} className={rank ? `rank-fill ${tierForRank(rank).color}` : ''} value={id} onChange={(e) => onChange(session.id, type, row, e.target.value)}><option value="">—</option>{members.map((candidate) => <option key={candidate.id} value={candidate.id}>{candidate.name}</option>)}</select> })}</div>)}</div>
}
