/* 端到端核对：明细列对齐、特殊清单口径、去重、空值、XLSX 数字与转义 */
import assert from 'node:assert'
import { BUILTIN_RULES } from '../src/logic/sizeRules'
import { runMerge, buildSummary, isSpecialPerson } from '../src/logic/merge'
import {
  DETAIL_HEADER,
  SPECIAL_HEADER,
  detailRows,
  specialRows,
  specialPersons,
  detailWorkbookSheets,
  orderWorkbookSheets,
  buildOrderSheet
} from '../src/logic/exporter'
import { buildXlsxBlob } from '../src/logic/xlsx'
import { toCsvText } from '../src/logic/csv'
import type { Person, Project } from '../src/logic/types'

const rule = BUILTIN_RULES[0]
let seq = 0
function makePerson(partial: Partial<Person>): Person {
  seq += 1
  return {
    id: `p${seq}`,
    name: `测试${seq}`,
    gender: 'male',
    orgUnit: '高一&(1)班',
    batch: '秋装',
    heightCm: 170,
    weightKg: 60,
    chestCm: 88,
    waistCm: 72,
    specialFlag: null,
    note: '',
    status: 'active',
    statusReason: '',
    anomaly: [],
    needsConfirm: false,
    possibleDuplicateOf: null,
    sourceRow: seq,
    source: 'import',
    result: null,
    createdAt: 0,
    ...partial
  }
}

const persons: Person[] = [
  makePerson({ name: '张三', heightCm: 170, chestCm: 88, waistCm: 72 }), // 常规 170/88A
  makePerson({
    name: '李四',
    orgUnit: '高二<2>班',
    note: '含 & < > " 引号',
    result: {
      sizeCode: '175/92A',
      ruleSizeCode: '170/88A',
      fit: 'A',
      ruleVersion: rule.version,
      manualOverride: { sizeCode: '175/92A', by: '王老师', reason: '肩宽加一档 & 确认', at: 1 }
    }
  }),
  makePerson({ name: '王五', specialFlag: 'PLUS' }), // 有效特殊
  makePerson({ name: '赵六', specialFlag: 'TALL', status: 'invalid', statusReason: '身高异常' }), // 有标记但无效
  makePerson({ name: '孙七', specialFlag: 'CUSTOM', status: 'duplicate', statusReason: '重复' }), // 有标记但重复排除
  makePerson({ name: '周八', heightCm: 168, chestCm: 90, waistCm: 70, note: '   ' }), // 纯空格备注
  makePerson({ name: '吴九', heightCm: 172, chestCm: 86, waistCm: 74, note: '' }) // 常规
]
// 人为重复同一 id（模拟同一条记录被取两遍），uniquePersons 必须兜底
persons.push(persons[0])

const project: Project = {
  id: 'prj1',
  name: '测试 <学校> & "一中"',
  kind: 'school',
  ruleVersion: rule.version,
  batches: ['秋装'],
  persons,
  imports: [],
  createdAt: 0,
  updatedAt: 0
}

runMerge(project, rule)
const summary = buildSummary(project, rule)

/* 1. 明细：表头 18 列，每行 18 列，关键列内容对位 */
const details = detailRows({ project, rule })
assert.strictEqual(details[0], DETAIL_HEADER)
assert.strictEqual(DETAIL_HEADER.length, 18)
details.slice(1).forEach((row, i) => {
  assert.strictEqual(row.length, DETAIL_HEADER.length, `明细第 ${i + 1} 行列数 ${row.length} != 表头 ${DETAIL_HEADER.length}`)
})
const liSi = details.find((r) => r[1] === '李四')!
const col = (title: string) => DETAIL_HEADER.indexOf(title)
assert.strictEqual(liSi[col('胸腰差(cm)')], '16')
assert.strictEqual(liSi[col('规则号型')], '170/88A')
assert.strictEqual(liSi[col('生效号型')], '175/92A')
assert.strictEqual(liSi[col('是否覆写')], '是')
assert.strictEqual(liSi[col('覆写人')], '王老师')
assert.strictEqual(liSi[col('覆写原因')], '肩宽加一档 & 确认')
assert.strictEqual(liSi[col('特殊体型')], '')
assert.strictEqual(liSi[col('状态')], '有效')
const wangWu = details.find((r) => r[1] === '王五')!
assert.strictEqual(wangWu[col('特殊体型')], '加肥加大')
const zhangSan = details.find((r) => r[1] === '张三')!
assert.strictEqual(zhangSan[col('是否覆写')], '否')
assert.strictEqual(zhangSan[col('规则号型')], '170/88A')

// 重复 id 只出现一次
assert.strictEqual(details.filter((r) => r[1] === '张三').length, 1, '同一记录不得出现两遍')
assert.strictEqual(details.length - 1, 7)

/* 2. 纯空格备注是有值的（与空备注区分），且保留原空格 */
const zhouBa = details.find((r) => r[1] === '周八')!
assert.strictEqual(zhouBa[col('备注')], '   ', '纯空格单元格必须保留为空格')
const wuJiu = details.find((r) => r[1] === '吴九')!
assert.strictEqual(wuJiu[col('备注')], '', '真空值才是空串')

/* 3. 特殊清单口径：仅有效 + 带标记（王五）；赵六/孙七不算 */
const specials = specialRows({ project, rule })
assert.strictEqual(specials[0], SPECIAL_HEADER)
assert.strictEqual(SPECIAL_HEADER.length, 13)
for (const row of specials.slice(1)) {
  assert.strictEqual(row.length, SPECIAL_HEADER.length, '清单行列数与表头不一致')
}
const specialNames = specials.slice(1).map((r) => r[1])
assert.deepStrictEqual(specialNames, ['王五'])
assert.strictEqual(specialPersons({ project, rule }).length, 1)
assert.strictEqual(specialPersons({ project, rule })[0].name, '王五')

// 与归并页 / 汇总完全同口径
assert.strictEqual(summary.totals.specialPersonCount, 1)
assert.strictEqual(summary.totals.specialQty, 1)
assert.ok(!isSpecialPerson(persons[3]) && !isSpecialPerson(persons[4]))
assert.strictEqual(summary.totals.validRows, 6)
assert.strictEqual(summary.totals.invalidRows, 1)
assert.strictEqual(summary.totals.duplicateRows, 1)
assert.strictEqual(summary.totals.accountedQty, 6)
assert.strictEqual(summary.conserved, true)

/* 4. XLSX：数量列为数值单元格；文本含 & < > " 时转义且保留空格 */
const ctx = { project, rule, summary, operator: 'tester', generatedAt: new Date() }
const sheets = detailWorkbookSheets({ project, rule })
assert.strictEqual(sheets.length, 2)

async function asText(blob: Blob): Promise<string> {
  const buffer = await blob.arrayBuffer()
  // stored zip：直接在字节里找 XML 片段
  return new TextDecoder().decode(new Uint8Array(buffer))
}

const blob = buildXlsxBlob(detailWorkbookSheets({ project, rule }))
const text = await asText(blob)
assert.ok(text.includes('肩宽加一档 &amp; 确认'), '& 必须转义')
assert.ok(text.includes('含 &amp; &lt; &gt; &quot; 引号'), '< > " 必须转义')
assert.ok(text.includes('<t xml:space="preserve">   </t>'), '纯空格单元格要保留并标 preserve')
assert.ok(!/<is><t>[^<]*[&<>][^<]*<\/t><\/is>/.test(text), '不得出现未转义的裸 XML 字符')

// 下单表数量：qty 是 number -> 数值单元格 <c r="E.."><v>1</v></c>，不带 inlineStr
const orderBlob = buildXlsxBlob(orderWorkbookSheets(ctx))
const orderText = await asText(orderBlob)
assert.ok(/<c r="E\d+"><v>\d+<\/v><\/c>/.test(orderText), '数量必须是数值单元格')
assert.ok(!/t="inlineStr"><is><t>\d+<\/t><\/is>/.test(orderText.match(/<c r="E\d+"[^>]*>.*?<\/c>/g)?.join('') ?? ''),
  '数量列不能写成文字')

const order = buildOrderSheet(ctx)
assert.strictEqual(order.totalQty, 6)
assert.ok(order.items.every((it) => typeof it.qty === 'number'))

/* 5. CSV：数字保持裸值可求和；含特殊字符的字段加引号；纯空格字段加引号 */
const csv = toCsvText(specials)
assert.ok(csv.includes('高一&(1)班'), '& 在 CSV 中不需要转义但应原样保留')
const detailCsv = toCsvText(details)
assert.ok(detailCsv.includes('"   "'), '纯空格字段 CSV 中应加引号保留')
assert.ok(detailCsv.includes('"含 & < > "" 引号"'), '含引号/特殊字符的字段 CSV 加引号转义')

console.log('全部断言通过 ✓')
console.log(`明细 ${details.length - 1} 行 × ${DETAIL_HEADER.length} 列；特殊清单 ${specials.length - 1} 人`)
