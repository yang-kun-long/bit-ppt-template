#!/usr/bin/env node
import PptxGenJS from 'pptxgenjs'

const pres = new PptxGenJS()
const slide = pres.addSlide()

// 测试各种弯曲箭头类型
const arrowTypes = [
  { type: 'curvedRightArrow', label: 'curvedRightArrow', x: 1, y: 1 },
  { type: 'curvedLeftArrow', label: 'curvedLeftArrow', x: 4, y: 1 },
  { type: 'curvedUpArrow', label: 'curvedUpArrow', x: 7, y: 1 },
  { type: 'curvedDownArrow', label: 'curvedDownArrow', x: 1, y: 3 },
  { type: 'bentArrow', label: 'bentArrow', x: 4, y: 3 },
  { type: 'bentUpArrow', label: 'bentUpArrow', x: 7, y: 3 },
  { type: 'uturnArrow', label: 'uturnArrow', x: 1, y: 5 },
  { type: 'circularArrow', label: 'circularArrow', x: 4, y: 5 },
]

arrowTypes.forEach(({ type, label, x, y }) => {
  // 添加箭头形状
  slide.addShape(pres.ShapeType[type], {
    x,
    y,
    w: 2,
    h: 1,
    fill: { color: '006C39' },
  })

  // 添加标签
  slide.addText(label, {
    x,
    y: y + 1.2,
    w: 2,
    h: 0.3,
    fontSize: 10,
    align: 'center',
  })
})

slide.addText('PptxGenJS 弯曲箭头支持测试', {
  x: 0.5,
  y: 0.3,
  w: 9,
  h: 0.5,
  fontSize: 20,
  bold: true,
  align: 'center',
})

await pres.writeFile({ fileName: 'output/curved-arrows-test.pptx' })
console.log('✓ 生成成功: output/curved-arrows-test.pptx')
