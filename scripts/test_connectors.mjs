#!/usr/bin/env node
import PptxGenJS from 'pptxgenjs'

const pres = new PptxGenJS()
const slide = pres.addSlide()

slide.addText('PowerPoint Connector 测试', {
  x: 0.5,
  y: 0.3,
  w: 9,
  h: 0.5,
  fontSize: 20,
  bold: true,
  align: 'center',
})

// 创建两个形状
const shape1 = slide.addShape(pres.ShapeType.rect, {
  x: 1,
  y: 2,
  w: 2,
  h: 1,
  fill: { color: '006C39' },
})

const shape2 = slide.addShape(pres.ShapeType.rect, {
  x: 6,
  y: 4,
  w: 2,
  h: 1,
  fill: { color: 'A13F3D' },
})

// 尝试添加连接线（如果 pptxgenjs 支持的话）
// 注意：pptxgenjs 可能不支持真正的 connector，只支持普通的 line
try {
  slide.addShape(pres.ShapeType.line, {
    x: 3,
    y: 2.5,
    w: 3,
    h: 1.5,
    line: {
      color: '000000',
      width: 2,
      endArrowType: 'arrow',
    },
  })
  console.log('✓ 添加了普通 line（带箭头）')
} catch (err) {
  console.log('✗ line 失败:', err.message)
}

// 测试是否有 addConnector 方法
if (typeof slide.addConnector === 'function') {
  console.log('✓ pptxgenjs 支持 addConnector')
  try {
    slide.addConnector({
      type: 'elbow',
      from: { x: 3, y: 2.5 },
      to: { x: 6, y: 4.5 },
      line: { color: '0000FF', width: 2 },
    })
  } catch (err) {
    console.log('✗ addConnector 调用失败:', err.message)
  }
} else {
  console.log('✗ pptxgenjs 不支持 addConnector 方法')
}

await pres.writeFile({ fileName: 'output/connectors-test.pptx' })
console.log('✓ 生成成功: output/connectors-test.pptx')
