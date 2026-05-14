#!/usr/bin/env node
import { readFile } from 'fs/promises'

const slide1 = await readFile('tmp/inspect-rag-ppt/ppt/slides/slide1.xml', 'utf-8')

// 提取所有形状的类型和位置
const shapes = slide1.match(/<p:sp[^>]*>[\s\S]*?<\/p:sp>/g) || []

console.log('\n=== 箭头和线条分析 ===\n')

let arrowCount = 0
let lineCount = 0
let otherCount = 0

shapes.forEach((shape, i) => {
  const prst = shape.match(/<a:prstGeom prst="([^"]+)"/)
  const xfrm = shape.match(/<a:xfrm[^>]*>([\s\S]*?)<\/a:xfrm>/)

  if (prst) {
    const type = prst[1]

    // 提取位置和尺寸
    let x = 0, y = 0, w = 0, h = 0
    if (xfrm) {
      const offMatch = xfrm[1].match(/<a:off x="(\d+)" y="(\d+)"/)
      const extMatch = xfrm[1].match(/<a:ext cx="(\d+)" cy="(\d+)"/)
      if (offMatch) {
        x = parseInt(offMatch[1]) / 914400 // EMU to inches
        y = parseInt(offMatch[2]) / 914400
      }
      if (extMatch) {
        w = parseInt(extMatch[1]) / 914400
        h = parseInt(extMatch[2]) / 914400
      }
    }

    if (type.includes('Arrow') || type.includes('arrow')) {
      arrowCount++
      if (arrowCount <= 5) {
        console.log(`箭头 ${arrowCount}: ${type}`)
        console.log(`  位置: (${x.toFixed(2)}, ${y.toFixed(2)})`)
        console.log(`  尺寸: ${w.toFixed(2)} x ${h.toFixed(2)} 英寸`)

        // 检查是否有旋转
        const rot = shape.match(/<a:xfrm[^>]*rot="(\d+)"/)
        if (rot) {
          const degrees = parseInt(rot[1]) / 60000
          console.log(`  旋转: ${degrees}°`)
        }
        console.log()
      }
    } else if (type.includes('line') || type.includes('Line') || type.includes('Connector')) {
      lineCount++
    } else {
      otherCount++
    }
  }
})

console.log(`\n=== 统计 ===`)
console.log(`箭头形状: ${arrowCount}`)
console.log(`直线/连接线: ${lineCount}`)
console.log(`其他形状: ${otherCount}`)
console.log(`总计: ${shapes.length}`)

console.log(`\n结论: 这个 PPT 用了 ${arrowCount} 个箭头形状 + ${lineCount} 条直线拼接`)
console.log(`问题: 每个箭头都是独立形状，移动框图时箭头不会自动跟随`)
