// PowerPoint Connector 注入器
// 在 pptxgenjs 生成的 PPTX 中注入真正的连接线

import JSZip from 'jszip'
import { readFile, writeFile } from 'fs/promises'
import { generateConnectorXML } from './connector-template.mjs'

/**
 * 在已生成的 PPTX 中注入 Connector
 * @param {string} inputPath - 输入 PPTX 路径
 * @param {string} outputPath - 输出 PPTX 路径
 * @param {Array} connectors - 连接线配置数组
 * @param {number} slideIndex - 幻灯片索引 (1-based)
 */
export async function injectConnectors(inputPath, outputPath, connectors, slideIndex = 1) {
  // 读取 PPTX
  const data = await readFile(inputPath)
  const zip = await JSZip.loadAsync(data)

  // 读取目标幻灯片 XML
  const slidePath = `ppt/slides/slide${slideIndex}.xml`
  const slideXML = await zip.file(slidePath).async('string')

  // 解析 XML，找到最大的形状 ID
  const idMatches = slideXML.match(/id="(\d+)"/g) || []
  let maxId = 0
  idMatches.forEach(match => {
    const id = parseInt(match.match(/\d+/)[0])
    if (id > maxId) maxId = id
  })

  // 如果 connector 包含位置信息而不是 shapeId，需要查找对应的形状
  const resolvedConnectors = connectors.map(conn => {
    if (conn.from.shapeId && conn.to.shapeId) {
      // 已经有 shapeId，直接使用
      return conn
    }

    // 根据位置查找形状 ID
    const fromShapeId = findShapeIdByPosition(slideXML, conn.from.x, conn.from.y)
    const toShapeId = findShapeIdByPosition(slideXML, conn.to.x, conn.to.y)

    if (!fromShapeId || !toShapeId) {
      console.warn(`警告: 无法找到位置 (${conn.from.x}, ${conn.from.y}) 或 (${conn.to.x}, ${conn.to.y}) 的形状`)
      return null
    }

    return {
      ...conn,
      from: { ...conn.from, shapeId: fromShapeId },
      to: { ...conn.to, shapeId: toShapeId },
    }
  }).filter(c => c !== null)

  // 生成 connector XML 片段
  const connectorFragments = resolvedConnectors.map((conn, index) => {
    return generateConnectorXML({
      id: maxId + index + 1,
      name: conn.name || `Connector ${index + 1}`,
      fromShapeId: conn.from.shapeId,
      toShapeId: conn.to.shapeId,
      fromConnectionSite: conn.from.site || 2,
      toConnectionSite: conn.to.site || 0,
      type: conn.type || 'elbow',
      color: conn.color || '000000',
      width: conn.width || 25400,
      arrowType: conn.arrowType || 'triangle',
    })
  })

  // 将 connector 插入到 </p:spTree> 之前
  const insertPosition = slideXML.lastIndexOf('</p:spTree>')
  if (insertPosition === -1) {
    throw new Error('无法找到 </p:spTree> 标签')
  }

  const modifiedXML =
    slideXML.slice(0, insertPosition) +
    connectorFragments.join('\n') +
    slideXML.slice(insertPosition)

  // 更新 ZIP
  zip.file(slidePath, modifiedXML)

  // 写入新文件
  const output = await zip.generateAsync({ type: 'nodebuffer' })
  await writeFile(outputPath, output)
}

/**
 * 根据位置查找形状 ID
 * @param {string} xml - 幻灯片 XML
 * @param {number} x - X 坐标（英寸）
 * @param {number} y - Y 坐标（英寸）
 * @returns {number|null} 形状 ID
 */
function findShapeIdByPosition(xml, x, y) {
  // 将英寸转换为 EMU (1 inch = 914400 EMU)
  const targetX = Math.round(x * 914400)
  const targetY = Math.round(y * 914400)
  const tolerance = 10000 // 允许 10000 EMU 的误差

  // 查找所有形状
  const shapePattern = /<p:sp>[\s\S]*?<\/p:sp>/g
  const shapes = xml.match(shapePattern) || []

  for (const shape of shapes) {
    // 提取形状 ID
    const idMatch = shape.match(/<p:cNvPr id="(\d+)"/)
    if (!idMatch) continue
    const shapeId = parseInt(idMatch[1])

    // 提取位置
    const xMatch = shape.match(/<a:off x="(\d+)"/)
    const yMatch = shape.match(/y="(\d+)"/)
    if (!xMatch || !yMatch) continue

    const shapeX = parseInt(xMatch[1])
    const shapeY = parseInt(yMatch[1])

    // 检查位置是否匹配
    if (Math.abs(shapeX - targetX) < tolerance && Math.abs(shapeY - targetY) < tolerance) {
      return shapeId
    }
  }

  return null
}

/**
 * 从 YAML 配置中提取 connector 信息
 * 支持在 slide 中定义 connectors 字段
 *
 * @example
 * slides:
 *   - layout: custom
 *     shapes:
 *       - id: box1
 *         type: rect
 *         x: 1
 *         y: 2
 *       - id: box2
 *         type: rect
 *         x: 5
 *         y: 4
 *     connectors:
 *       - from: { shape: box1, site: 2 }
 *         to: { shape: box2, site: 0 }
 *         type: elbow
 *         color: "000000"
 */
export function extractConnectorsFromSlide(slideData, shapeIdMap) {
  if (!slideData.connectors || !Array.isArray(slideData.connectors)) {
    return []
  }

  return slideData.connectors.map(conn => {
    // 将形状名称映射到实际的形状 ID
    const fromShapeId = shapeIdMap[conn.from.shape]
    const toShapeId = shapeIdMap[conn.to.shape]

    if (!fromShapeId || !toShapeId) {
      throw new Error(`无法找到形状: ${conn.from.shape} 或 ${conn.to.shape}`)
    }

    return {
      from: {
        shapeId: fromShapeId,
        site: conn.from.site || 2,
      },
      to: {
        shapeId: toShapeId,
        site: conn.to.site || 0,
      },
      type: conn.type || 'elbow',
      color: conn.color || '000000',
      width: conn.width || 25400,
      arrowType: conn.arrowType || 'triangle',
      name: conn.name,
    }
  })
}
