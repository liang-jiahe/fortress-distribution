import { useEffect, useMemo, useRef, useState } from 'react'
import ExcelJS from 'exceljs'

type Member = {
  id: string
  name: string
  power: number
  weeklyPower: number
  score: number | null
  remark: string
  order: number
}

type PackageType = 'fire' | 'mid1' | 'mid2'
type Schedule = Record<string, string | null>
type AccessoryName = '手镯' | '戒指' | '耳环' | '腰带' | '项链' | '徽章'
type QueueEntry = { id: string; name: string; addedAt: string }

type Session = {
  id: string
  label: string
  block: 'pink' | 'orange' | 'yellow' | 'green' | 'cyan' | 'blue'
  startColumn: number
  startRow: number
}

const PACKAGE_LABELS: Record<PackageType, string> = { fire: '火', mid1: '中一', mid2: '中二' }
const TYPE_ORDER: PackageType[] = ['fire', 'mid1', 'mid2']
const ACCESSORIES: { name: AccessoryName; icon: string; color: string }[] = [
  { name: '手镯', icon: '🪬', color: 'pink' },
  { name: '戒指', icon: '💍', color: 'yellow' },
  { name: '耳环', icon: '✨', color: 'cyan' },
  { name: '腰带', icon: '🎀', color: 'green' },
  { name: '项链', icon: '📿', color: 'orange' },
  { name: '徽章', icon: '🏵️', color: 'blue' },
]
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

function emptyQueues(): Record<AccessoryName, QueueEntry[]> {
  return { 手镯: [], 戒指: [], 耳环: [], 腰带: [], 项链: [], 徽章: [] }
}

function sweepAccessoryQueues(source: Record<AccessoryName, QueueEntry[]>) {
  return ACCESSORIES.reduce((next, accessory) => {
    next[accessory.name] = source[accessory.name].slice(1)
    return next
  }, emptyQueues())
}

function sundayDateKey() {
  const now = new Date()
  return `${now.getFullYear()}-${now.getMonth() + 1}-${now.getDate()}`
}

const SAMPLE_NAMES = ['太初星', '关注塔菲喵', '花云青', '无双', '御茨星', '念君夏', '夏弥', '谦灵星', '小星星', '沈七涵', '晓行星', '心芯星', '猫猫星', '绝地', '鸿鹄', '云岫', '亿丈龙我', '椛七', '超级萝卜大王', '拳王', '浅帐星', '别急稳一手', '小苏在这里', '时愿星', '伦敦街尾吻别', '33', '季时雨花知否', '我一直都在', '弦', '白慕']
const SAMPLE_POWER = [3852, 3751, 3736, 3258, 3220, 3077, 2867, 2881, 2802, 2746, 2732, 2684, 2633, 2567, 2547, 2542, 2469, 2460, 2408, 2405, 2323, 2314, 2306, 2294, 2218, 2206, 2193, 2189, 2110, 2096]
const SAMPLE_SCORE_BY_NAME: Record<string, number> = {
  '太初星': 36, '关注塔菲喵': 36, '花云青': 36, '无双': 36, '御茨星': 35, '念君夏': 36, '夏弥': 36, '谦灵星': 36, '小星星': 35, '沈七涵': 16, '晓行星': 36, '心芯星': 36, '猫猫星': 36, '绝地': 36, '鸿鹄': 26, '云岫': 26, '亿丈龙我': 36, '椛七': 36, '超级萝卜大王': 5, '拳王': 36, '浅帐星': 36, '别急稳一手': 36, '小苏在这里': 15, '时愿星': 17, '伦敦街尾吻别': 36, '33': 36, '季时雨花知否': 36, '我一直都在': 15, '弦': 26, '白慕': 36,
}

function makeSampleMembers(): Member[] {
  return SAMPLE_NAMES.map((name, index) => ({ id: `m-${index + 1}`, name, power: SAMPLE_POWER[index], weeklyPower: 0, score: SAMPLE_SCORE_BY_NAME[name] ?? null, remark: '', order: index }))
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
  return [...byName.entries()].map(([name, data], index) => ({ id: `m-${Date.now()}-${index}`, name, power: data.power ?? 0, weeklyPower: 0, score: data.score ?? null, remark: data.remark ?? '', order: data.order ?? index })).sort((a, b) => a.order - b.order)
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
  sheet.mergeCells('A28:P28'); sheet.getCell('A28').value = '署名：繁星の猫猫星'; sheet.getCell('A28').alignment = center; sheet.getCell('A28').font = { name: '楷体', size: 12, italic: true, color: { argb: 'FF1B6B4F' } }
  const buffer = await workbook.xlsx.writeBuffer(); const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); const url = URL.createObjectURL(blob); const link = document.createElement('a'); link.href = url; link.download = '繁星本周要塞包分配.xlsx'; link.click(); URL.revokeObjectURL(url)
}

export default function App() {
  const [members, setMembers] = useState<Member[]>(() => { try { const saved = localStorage.getItem('fortress-members'); return saved ? JSON.parse(saved) : makeSampleMembers() } catch { return makeSampleMembers() } })
  const [contest, setContest] = useState(false); const [search, setSearch] = useState(''); const [notice, setNotice] = useState('已加载示例数据，可直接编辑或导入本周表格。'); const [activeSection, setActiveSection] = useState('matrix'); const fileRef = useRef<HTMLInputElement>(null)
  const [queues, setQueues] = useState<Record<AccessoryName, QueueEntry[]>>(() => { try { const saved = localStorage.getItem('fortress-accessory-queues'); return saved ? { ...emptyQueues(), ...JSON.parse(saved) } : emptyQueues() } catch { return emptyQueues() } })
  const [queueInputs, setQueueInputs] = useState<Record<AccessoryName, string>>(() => ({ 手镯: '', 戒指: '', 耳环: '', 腰带: '', 项链: '', 徽章: '' }))
  const [lastSweep, setLastSweep] = useState(() => localStorage.getItem('fortress-accessory-last-sweep') || '')
  const ranked = useMemo(() => rankMembers(members), [members]); const powerRanked = useMemo(() => powerRankMembers(members), [members]); const schedule = useMemo(() => buildAutoSchedule(ranked), [ranked]); const scoreMax = contest ? 57 : 37
  const weeklyPowerTotal = useMemo(() => members.reduce((total, member) => total + (member.weeklyPower || 0), 0), [members])
  const counts = useMemo(() => { const result = new Map<string, { fire: number; middle: number }>(); Object.entries(schedule).forEach(([key, id]) => { if (!id) return; const type = key.split(':')[1] as PackageType; const current = result.get(id) ?? { fire: 0, middle: 0 }; if (type === 'fire') current.fire += 1; else current.middle += 1; result.set(id, current) }); return result }, [schedule])
  useEffect(() => { localStorage.setItem('fortress-members', JSON.stringify(members)) }, [members])
  useEffect(() => { localStorage.setItem('fortress-accessory-queues', JSON.stringify(queues)) }, [queues])
  useEffect(() => { if (new Date().getDay() === 0 && lastSweep !== sundayDateKey()) { setQueues((current) => sweepAccessoryQueues(current)); setLastSweep(sundayDateKey()); localStorage.setItem('fortress-accessory-last-sweep', sundayDateKey()); setNotice('今天是周日，已为每种饰品发放队首名额。喵～') } }, [lastSweep])
  const updateMember = (id: string, patch: Partial<Member>) => { setMembers((current) => current.map((member) => member.id === id ? { ...member, ...patch } : member)) }
  const addMember = () => { const index = members.length + 1; setMembers((current) => [...current, { id: `m-${Date.now()}`, name: `新成员${index}`, power: 0, weeklyPower: 0, score: null, remark: '', order: current.length }]); setNotice('已新增成员，请填写姓名、战力和考核分。') }
  const removeMember = (id: string) => { setMembers((current) => current.filter((member) => member.id !== id)); setNotice('成员已删除，排名和矩阵已更新。') }
  const calculateWeeklyPower = () => { const increase = weeklyPowerTotal; setMembers((current) => current.map((member) => ({ ...member, power: member.power + (member.weeklyPower || 0), weeklyPower: 0 }))); setNotice(`本周战力已完成计算，总计增加 ${increase} 万。`) }
  const visibleMembers = powerRanked.filter((member) => member.name.toLowerCase().includes(search.toLowerCase()))
  const reset = () => { setMembers(makeSampleMembers()); setNotice('已恢复示例数据。') }
  const clear = () => { setMembers([]); setNotice('已清空成员数据。') }
  const addQueueEntry = (accessory: AccessoryName) => { const name = queueInputs[accessory].trim(); if (!name) { setNotice(`请先填写想要${accessory}的猫咪姓名。`); return } if (queues[accessory].some((entry) => entry.name.trim().toLowerCase() === name.toLowerCase())) { setNotice(`${name} 已经在${accessory}队列里啦。`); return } setQueues((current) => ({ ...current, [accessory]: [...current[accessory], { id: `q-${Date.now()}-${accessory}`, name, addedAt: new Date().toISOString() }] })); setQueueInputs((current) => ({ ...current, [accessory]: '' })); setNotice(`${name} 已加入${accessory}排队。`) }
  const removeQueueEntry = (accessory: AccessoryName, id: string) => { setQueues((current) => ({ ...current, [accessory]: current[accessory].filter((entry) => entry.id !== id) })); setNotice('已标记为分发完成，队列已更新。') }
  const sweepQueuesNow = () => { const waiting = ACCESSORIES.reduce((total, accessory) => total + (queues[accessory.name][0] ? 1 : 0), 0); setQueues((current) => sweepAccessoryQueues(current)); setLastSweep(sundayDateKey()); localStorage.setItem('fortress-accessory-last-sweep', sundayDateKey()); setNotice(`已手动发放 ${waiting} 个饰品队首名额。`) }
  const handleImport = async (event: React.ChangeEvent<HTMLInputElement>) => { const file = event.target.files?.[0]; if (!file) return; try { const imported = await importWorkbook(file); if (!imported.length) throw new Error('没有识别到成员'); setMembers(imported); setNotice(`已导入 ${imported.length} 名成员。`) } catch (error) { setNotice(`导入失败：${error instanceof Error ? error.message : '文件格式不正确'}`) } finally { event.target.value = '' } }
  const goTo = (section: string) => { setActiveSection(section); document.getElementById(section)?.scrollIntoView({ behavior: 'smooth', block: 'start' }) }
  return <div className="app-shell">
    <span className="cat-sticker sticker-cat" aria-hidden="true">🐈‍⬛</span><span className="cat-sticker sticker-paw" aria-hidden="true">🐾</span><span className="cat-sticker sticker-heart" aria-hidden="true">😻</span>
    <aside className="sidebar"><div className="brand-mark">🐱</div><h1>繁星要塞</h1><nav><button className={activeSection === 'matrix' ? 'active' : ''} onClick={() => goTo('matrix')}>🏰 要塞分包</button><button className={activeSection === 'members' ? 'active' : ''} onClick={() => goTo('members')}>🐾 成员管理</button><button className={activeSection === 'ranking' ? 'active' : ''} onClick={() => goTo('members')}>🏆 排名与考核</button><button className={activeSection === 'accessories' ? 'active' : ''} onClick={() => goTo('accessories')}>🎁 饰品排队</button><button className={activeSection === 'instructions' ? 'active' : ''} onClick={() => goTo('instructions')}>📜 使用说明</button></nav><div className="sidebar-note">数据只保存在当前浏览器<br />导入即用，导出即走</div></aside>
    <main className="content">
      <header className="topbar"><div><div className="eyebrow">FORTRESS DISTRIBUTION</div><h2>🐱 本周要塞包分配</h2></div><div className="toolbar"><button className="btn secondary" onClick={() => fileRef.current?.click()}>🐾 导入 XLSX</button><input ref={fileRef} type="file" accept=".xlsx,.xls" hidden onChange={handleImport} /><button className="btn primary" onClick={() => exportWorkbook(members, schedule, ranked)}>🐟 导出单表 XLSX</button></div></header>
      <section className="notice">{notice}<span className="notice-right"><label className="toggle"><input type="checkbox" checked={contest} onChange={(e) => setContest(e.target.checked)} /><span></span> 争霸周（最高 {scoreMax} 分）</label></span></section>
      <section className="stats"><div className="stat-card"><span>成员人数</span><strong>{members.length}</strong><small>{members.length === 30 ? '模板完整' : '目标 30 人'}</small></div><div className="stat-card"><span>当前最高分</span><strong>{ranked[0]?.score ?? '—'}</strong><small>普通周 37 · 争霸周 57</small></div><div className="stat-card"><span>本周提升</span><strong>{weeklyPowerTotal}</strong><small>输入后点击“计算完成”</small></div><div className="stat-card"><span>火/中包总量</span><strong>40 / 80</strong><small>8 个时段完整分配</small></div></section>
      <section id="matrix" className="panel matrix-panel"><div className="panel-heading"><div><span className="eyebrow">AUTO LAYOUT</span><h3>😸 彩色分包矩阵</h3></div><button className="btn ghost" onClick={() => setNotice('矩阵会根据当前考核分排名自动生成。')}>🐾 重新生成矩阵</button></div><div className="matrix-scroll"><div className="matrix-grid">{SESSIONS.map((session) => <SessionBlock key={session.id} session={session} schedule={schedule} ranked={ranked} members={members} />)}</div></div><div className="legend">{TIERS.map((tier) => <span key={tier.min}><i className={`swatch ${tier.color}`}></i>{tier.min}-{tier.max} 名</span>)}</div></section>
      <section id="members" className="panel"><span id="ranking" className="anchor-target"></span><div className="panel-heading"><div><span className="eyebrow">MEMBER ROSTER</span><h3>成员与考核分</h3></div><div className="row-actions"><input className="search" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="搜索姓名" /><button className="btn primary" onClick={addMember}>＋ 新增成员</button><button className="btn ghost" onClick={calculateWeeklyPower}>✓ 计算完成</button><button className="btn ghost" onClick={reset}>恢复示例</button><button className="btn danger" onClick={clear}>清空</button></div></div><div className="table-wrap"><table><thead><tr><th>战力排名</th><th>成员姓名</th><th>总战力</th><th>本周提升</th><th>分包排名</th><th>考核分</th><th>火/中包</th><th>币数</th><th>备注</th><th></th></tr></thead><tbody>{visibleMembers.map((member) => { const powerRank = powerRanked.findIndex((entry) => entry.id === member.id) + 1; const packageRank = ranked.findIndex((entry) => entry.id === member.id) + 1; const tier = packageRank ? tierForRank(packageRank) : TIERS[0]; const count = counts.get(member.id) ?? { fire: 0, middle: 0 }; return <tr key={member.id}><td><span className="rank-pill">{powerRank || '—'}</span></td><td><input value={member.name} onChange={(e) => updateMember(member.id, { name: e.target.value })} /></td><td><input type="number" value={member.power} onChange={(e) => updateMember(member.id, { power: Number(e.target.value) || 0 })} /></td><td><input className="weekly-power" type="number" min="0" value={member.weeklyPower || ''} placeholder="+万" onChange={(e) => updateMember(member.id, { weeklyPower: Number(e.target.value) || 0 })} /></td><td><span className={`tier-dot ${tier.color}`}>{packageRank || '—'}</span></td><td><input type="number" min="0" max={scoreMax} value={member.score ?? ''} placeholder="未填" onChange={(e) => updateMember(member.id, { score: e.target.value === '' ? null : Number(e.target.value) })} /></td><td>{count.fire} 火 / {count.middle} 中</td><td>{packageRank ? tier.coins : '—'}</td><td><input value={member.remark} onChange={(e) => updateMember(member.id, { remark: e.target.value })} placeholder="备注" /></td><td><button className="icon-btn" title="删除成员" onClick={() => removeMember(member.id)}>×</button></td></tr> })}</tbody></table>{visibleMembers.length === 0 && <div className="empty">没有匹配的成员，点击“新增成员”开始。</div>}</div></section>
      <section id="instructions" className="panel rules-panel"><div><span className="eyebrow">PACKAGE RULES</span><h3>🐾 繁星猫咪分包协议</h3></div><div className="protocol-copy"><p>喵呜～每周六晚统计一次战力与考核分，按时参加活动的小猫咪基本不会被扣包。</p><ul><li>考核分越高，分包排名越靠前；同分时战力高者优先。</li><li>普通周满分 37 分；争霸周满分 57 分。</li><li>每周固定 8 个时段，共 40 个火包和 80 个中包。</li><li>1–5 名领取 2 火 3 中，26–30 名领取 5 中，其余档位按卡片执行。</li><li>特殊奖励或扣包请在备注里写清楚，保持猫咪指挥部记录整洁。</li></ul></div><div className="tier-cards">{TIERS.map((tier) => <div className={`tier-card ${tier.color}`} key={tier.min}><b>{tier.min}-{tier.max}</b><span>{tier.fire} 火 · {tier.middle} 中</span><strong>{tier.coins} 币</strong></div>)}</div><p className="rule-copy">本协议以互相提醒、按时参加、公开透明为原则：不抢包、不漏包，大家一起做一只守规矩的小猫咪。</p></section>
      <section id="accessories" className="panel accessory-panel"><div className="panel-heading"><div><span className="eyebrow">CAT ACCESSORY QUEUE</span><h3>🎁 饰品排队小站</h3></div><button className="btn ghost" onClick={sweepQueuesNow}>🐾 周日发放首位</button></div><p className="accessory-intro">想要哪件饰品，就在对应的小猫队伍里留下名字。每周日发放每种饰品的第一位，也可以随时手动标记“已分发”。</p><div className="accessory-grid">{ACCESSORIES.map((accessory) => <div className={`accessory-card ${accessory.color}`} key={accessory.name}><div className="accessory-title"><span className="accessory-icon">{accessory.icon}</span><div><strong>{accessory.name}</strong><small>{queues[accessory.name].length} 只猫咪排队</small></div></div><div className="queue-add"><input value={queueInputs[accessory.name]} onChange={(event) => setQueueInputs((current) => ({ ...current, [accessory.name]: event.target.value }))} onKeyDown={(event) => { if (event.key === 'Enter') addQueueEntry(accessory.name) }} placeholder="输入猫咪姓名" /><button className="cat-add" onClick={() => addQueueEntry(accessory.name)}>＋</button></div>{queues[accessory.name].length ? <ol className="queue-list">{queues[accessory.name].map((entry, index) => <li key={entry.id}><span className="queue-number">{index + 1}</span><span className="queue-name">{entry.name}</span><button onClick={() => removeQueueEntry(accessory.name, entry.id)}>已分发</button></li>)}</ol> : <div className="queue-empty">🐱 还没有猫咪排队</div>}</div>)}</div><p className="queue-note">自动规则：每周日打开系统时，每种饰品只会自动发放一次队首；队列数据保存在当前浏览器。</p></section>
      <footer className="signature">署名：繁星の猫猫星</footer>
    </main>
  </div>
}

function SessionBlock({ session, schedule, ranked, members }: { session: Session; schedule: Schedule; ranked: Member[]; members: Member[] }) {
  const memberById = new Map(members.map((member) => [member.id, member]))
  return <div className={`session-block ${session.block}`}><div className="session-title">🐾 {session.label}</div><div className="session-types">{TYPE_ORDER.map((type) => { const id = schedule[cellKey(session.id, type, 0)] ?? ''; const member = memberById.get(id); const rank = member ? ranked.findIndex((entry) => entry.id === member.id) + 1 : 0; return <div className={rank ? `header-fill ${tierForRank(rank).color}` : ''} key={type}>{PACKAGE_LABELS[type]}</div> })}</div>{Array.from({ length: 5 }, (_, row) => <div className="session-row" key={row}>{TYPE_ORDER.map((type) => { const id = schedule[cellKey(session.id, type, row)] ?? ''; const member = memberById.get(id); const rank = member ? ranked.findIndex((entry) => entry.id === member.id) + 1 : 0; return <div key={type} className={`matrix-name ${rank ? `rank-fill ${tierForRank(rank).color}` : ''}`}>{member?.name ?? '—'}</div> })}</div>)}</div>
}
