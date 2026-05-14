#!/usr/bin/env node
// 创建一个包含真正 Connector 的测试 PPT，用于研究 OpenXML 结构
import PptxGenJS from 'pptxgenjs'

const pres = new PptxGenJS()
const slide = pres.addSlide()

// 创建两个形状，稍后我们会手动在 PowerPoint 中添加 connector
const shape1 = slide.addShape(pres.ShapeType.rect, {
  x: 1,
  y: 2,
  w: 2,
  h: 1,
  fill: { color: '006C39' },
})

slide.addText('Shape 1', {
  x: 1,
  y: 2,
  w: 2,
  h: 1,
  align: 'center',
  valign: 'middle',
  color: 'FFFFFF',
})

const shape2 = slide.addShape(pres.ShapeType.rect, {
  x: 6,
  y: 4,
  w: 2,
  h: 1,
  fill: { color: 'A13F3D' },
})

slide.addText('Shape 2', {
  x: 6,
  y: 4,
  w: 2,
  h: 1,
  align: 'center',
  valign: 'middle',
  color: 'FFFFFF',
})

// 添加一条普通 line 作为对比
slide.addShape(pres.ShapeType.line, {
  x: 3,
  y: 2.5,
  w: 3,
  h: 1.5,
  line: {
    color: '000000',
    width: 2,
    endArrowType: 'triangle',
  },
})

await pres.writeFile({ fileName: 'output/before-connector.pptx' })
console.log('✓ 生成 output/before-connector.pptx')
console.log('请在 PowerPoint 中：')
console.log('1. 打开这个文件')
console.log('2. 插入 -> 形状 -> 连接线（Connector）')
console.log('3. 连接两个矩形')
console.log('4. 另存为 output/after-connector.pptx')
console.log('5. 然后运行 node scripts/extract_connector_xml.mjs')
