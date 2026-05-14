#!/usr/bin/env node
// 测试 Connector 注入功能

import PptxGenJS from 'pptxgenjs'
import { injectConnectors } from '../src/connectors.mjs'
import { mkdir } from 'fs/promises'

// 确保输出目录存在
await mkdir('output', { recursive: true })

// 第一步：用 pptxgenjs 生成基础 PPT
console.log('步骤 1: 生成基础 PPT...')
const pres = new PptxGenJS()
const slide = pres.addSlide()

// 创建三个形状
slide.addShape(pres.ShapeType.rect, {
  x: 1,
  y: 1,
  w: 2,
  h: 1,
  fill: { color: '006C39' },
})
slide.addText('数据源', {
  x: 1,
  y: 1,
  w: 2,
  h: 1,
  align: 'center',
  valign: 'middle',
  color: 'FFFFFF',
  fontSize: 18,
  bold: true,
})

slide.addShape(pres.ShapeType.rect, {
  x: 5,
  y: 1,
  w: 2,
  h: 1,
  fill: { color: '0066CC' },
})
slide.addText('处理层', {
  x: 5,
  y: 1,
  w: 2,
  h: 1,
  align: 'center',
  valign: 'middle',
  color: 'FFFFFF',
  fontSize: 18,
  bold: true,
})

slide.addShape(pres.ShapeType.rect, {
  x: 9,
  y: 1,
  w: 2,
  h: 1,
  fill: { color: 'A13F3D' },
})
slide.addText('输出', {
  x: 9,
  y: 1,
  w: 2,
  h: 1,
  align: 'center',
  valign: 'middle',
  color: 'FFFFFF',
  fontSize: 18,
  bold: true,
})

const basePath = 'output/connector-test-base.pptx'
await pres.writeFile({ fileName: basePath })
console.log(`✓ 生成基础 PPT: ${basePath}`)

// 第二步：注入 Connector
console.log('\n步骤 2: 注入 Connector...')

// 注意：pptxgenjs 生成的形状 ID 从 2 开始（1 是 spTree）
// 我们需要找到实际的形状 ID
const connectors = [
  {
    name: 'Flow 1',
    from: { shapeId: 2, site: 1 }, // 第一个矩形的右侧
    to: { shapeId: 4, site: 3 },   // 第二个矩形的左侧
    type: 'elbow',
    color: '000000',
    width: 19050, // 0.75pt
    arrowType: 'triangle',
  },
  {
    name: 'Flow 2',
    from: { shapeId: 4, site: 1 }, // 第二个矩形的右侧
    to: { shapeId: 6, site: 3 },   // 第三个矩形的左侧
    type: 'elbow',
    color: '000000',
    width: 19050,
    arrowType: 'triangle',
  },
]

const outputPath = 'output/connector-test-final.pptx'
await injectConnectors(basePath, outputPath, connectors, 1)
console.log(`✓ 注入 Connector: ${outputPath}`)

console.log('\n✓ 完成！请在 PowerPoint 中打开以下文件对比：')
console.log(`  - 基础版本: ${basePath}`)
console.log(`  - 注入版本: ${outputPath}`)
console.log('\n在注入版本中，尝试移动矩形，连接线应该自动跟随！')
