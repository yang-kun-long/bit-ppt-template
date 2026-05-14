#!/usr/bin/env node
// 验证生成的 PPT 中的 Connector

import JSZip from 'jszip'
import { readFile } from 'fs/promises'

const filePath = process.argv[2] || 'output/connector-demo.pptx'

console.log(`检查文件: ${filePath}\n`)

const data = await readFile(filePath)
const zip = await JSZip.loadAsync(data)

// 获取所有幻灯片
const slideFiles = Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))

console.log(`找到 ${slideFiles.length} 张幻灯片\n`)

for (const slideFile of slideFiles) {
  const slideXML = await zip.file(slideFile).async('string')
  const slideNum = slideFile.match(/slide(\d+)\.xml/)[1]

  // 检查是否包含 cxnSp 元素
  const cxnSpMatches = slideXML.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g)

  if (cxnSpMatches) {
    console.log(`📊 幻灯片 ${slideNum}: 找到 ${cxnSpMatches.length} 个 Connector`)

    cxnSpMatches.forEach((xml, index) => {
      const nameMatch = xml.match(/name="([^"]+)"/)
      const stCxnMatch = xml.match(/<p:stCxn id="(\d+)" idx="(\d+)"/)
      const endCxnMatch = xml.match(/<p:endCxn id="(\d+)" idx="(\d+)"/)

      if (nameMatch) console.log(`  - ${nameMatch[1]}`)
      if (stCxnMatch && endCxnMatch) {
        console.log(`    从形状 ${stCxnMatch[1]} (连接点 ${stCxnMatch[2]}) → 形状 ${endCxnMatch[1]} (连接点 ${endCxnMatch[2]})`)
      }
    })
    console.log()
  } else {
    console.log(`📄 幻灯片 ${slideNum}: 无 Connector (使用传统 line 形状)\n`)
  }
}

console.log('✓ 检查完成！')
