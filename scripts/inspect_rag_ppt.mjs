#!/usr/bin/env node
import { exec } from 'child_process'
import { promisify } from 'util'
import { readFile } from 'fs/promises'

const execAsync = promisify(exec)

const pptxPath = 'D:\\东南大学\\交叉\\reports\\2026-05-14\\ppt\\RAG_editable_architecture_2026-05-14.pptx'
const tempDir = 'tmp/inspect-rag-ppt'

// 解压 PPTX
await execAsync(`unzip -o "${pptxPath}" -d ${tempDir}`)

// 读取第一张幻灯片
const slide1 = await readFile(`${tempDir}/ppt/slides/slide1.xml`, 'utf-8')

// 查找所有 sp (shape) 和 cxnSp (connector) 元素
const shapes = slide1.match(/<p:sp[^>]*>[\s\S]*?<\/p:sp>/g) || []
const connectors = slide1.match(/<p:cxnSp[^>]*>[\s\S]*?<\/p:cxnSp>/g) || []

console.log(`\n=== RAG PPT 结构分析 ===`)
console.log(`形状数量: ${shapes.length}`)
console.log(`连接线数量: ${connectors.length}`)

// 分析连接线类型
if (connectors.length > 0) {
  console.log(`\n--- 连接线详情 ---`)
  connectors.slice(0, 3).forEach((conn, i) => {
    const prst = conn.match(/<a:prstGeom prst="([^"]+)"/)
    const beginCxn = conn.match(/<a:stCxn[^>]*id="(\d+)"/)
    const endCxn = conn.match(/<a:endCxn[^>]*id="(\d+)"/)
    console.log(`\n连接线 ${i + 1}:`)
    console.log(`  类型: ${prst ? prst[1] : '未知'}`)
    console.log(`  起点连接: ${beginCxn ? beginCxn[1] : '无'}`)
    console.log(`  终点连接: ${endCxn ? endCxn[1] : '无'}`)
  })
}

// 检查是否有 line 形状（非连接线）
const lines = shapes.filter(s => s.includes('prst="line"') || s.includes('prst="straightConnector'))
console.log(`\n普通直线数量: ${lines.length}`)

console.log(`\n结论: ${connectors.length > 0 ? '使用了真正的 Connector' : '只用了普通 Shape 拼接'}`)
