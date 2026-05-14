#!/usr/bin/env node
// 验证 Connector 是否正确注入

import JSZip from 'jszip'
import { readFile } from 'fs/promises'

const filePath = 'output/connector-test-final.pptx'

console.log(`检查文件: ${filePath}\n`)

const data = await readFile(filePath)
const zip = await JSZip.loadAsync(data)

// 读取第一张幻灯片
const slideXML = await zip.file('ppt/slides/slide1.xml').async('string')

// 检查是否包含 cxnSp 元素
const cxnSpMatches = slideXML.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g)

if (!cxnSpMatches) {
  console.log('❌ 未找到 <p:cxnSp> 元素')
  process.exit(1)
}

console.log(`✓ 找到 ${cxnSpMatches.length} 个 Connector\n`)

cxnSpMatches.forEach((xml, index) => {
  console.log(`--- Connector ${index + 1} ---`)

  // 提取关键信息
  const nameMatch = xml.match(/name="([^"]+)"/)
  const prstMatch = xml.match(/prst="([^"]+)"/)
  const stCxnMatch = xml.match(/<p:stCxn id="(\d+)" idx="(\d+)"/)
  const endCxnMatch = xml.match(/<p:endCxn id="(\d+)" idx="(\d+)"/)
  const colorMatch = xml.match(/<a:srgbClr val="([^"]+)"/)
  const arrowMatch = xml.match(/<a:tailEnd type="([^"]+)"/)

  if (nameMatch) console.log(`  名称: ${nameMatch[1]}`)
  if (prstMatch) console.log(`  类型: ${prstMatch[1]}`)
  if (stCxnMatch) console.log(`  起点: 形状 ${stCxnMatch[1]}, 连接点 ${stCxnMatch[2]}`)
  if (endCxnMatch) console.log(`  终点: 形状 ${endCxnMatch[1]}, 连接点 ${endCxnMatch[2]}`)
  if (colorMatch) console.log(`  颜色: #${colorMatch[1]}`)
  if (arrowMatch) console.log(`  箭头: ${arrowMatch[1]}`)
  console.log()
})

console.log('✓ Connector 注入成功！')
console.log('\n请在 PowerPoint 中打开文件，尝试移动矩形验证连接线是否自动跟随。')
