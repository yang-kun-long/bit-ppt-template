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

  // 生成 connector XML 片段
  const connectorFragments = connectors.map((conn, index) => {
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
