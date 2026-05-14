// Connector OpenXML 模板生成器
// 基于 Office Open XML 规范：https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_cxnSp_topic_ID0EQ2HPB.html

/**
 * 生成 PowerPoint Connector 的 OpenXML 片段
 * @param {Object} options
 * @param {number} options.id - 形状 ID
 * @param {string} options.name - 形状名称
 * @param {number} options.fromShapeId - 起点形状 ID
 * @param {number} options.toShapeId - 终点形状 ID
 * @param {number} options.fromConnectionSite - 起点连接点索引 (0-3: 上右下左)
 * @param {number} options.toConnectionSite - 终点连接点索引
 * @param {string} options.type - 连接线类型: 'straight' | 'elbow' | 'curved'
 * @param {string} options.color - 线条颜色 (RGB hex)
 * @param {number} options.width - 线条宽度 (EMU, 默认 25400 = 1pt)
 * @param {string} options.arrowType - 箭头类型: 'none' | 'triangle' | 'arrow'
 * @returns {string} OpenXML 片段
 */
export function generateConnectorXML(options) {
  const {
    id,
    name = `Connector ${id}`,
    fromShapeId,
    toShapeId,
    fromConnectionSite = 2, // 默认从下方连接
    toConnectionSite = 0,   // 默认到上方连接
    type = 'elbow',         // 默认折线
    color = '000000',
    width = 25400,          // 1pt
    arrowType = 'triangle',
  } = options

  // 连接线类型映射
  const connectorTypes = {
    straight: 'straightConnector1',
    elbow: 'bentConnector3',      // 一个拐角
    curved: 'curvedConnector3',   // 曲线
  }

  const prst = connectorTypes[type] || 'bentConnector3'

  return `<p:cxnSp>
  <p:nvCxnSpPr>
    <p:cNvPr id="${id}" name="${name}"/>
    <p:cNvCxnSpPr>
      <a:cxnSpLocks/>
    </p:cNvCxnSpPr>
    <p:nvPr/>
  </p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="0" cy="0"/>
    </a:xfrm>
    <a:prstGeom prst="${prst}">
      <a:avLst/>
    </a:prstGeom>
    <a:ln w="${width}">
      <a:solidFill>
        <a:srgbClr val="${color}"/>
      </a:solidFill>
      <a:headEnd type="none"/>
      <a:tailEnd type="${arrowType}"/>
    </a:ln>
  </p:spPr>
  <p:style>
    <a:lnRef idx="1">
      <a:schemeClr val="accent1"/>
    </a:lnRef>
    <a:fillRef idx="0">
      <a:schemeClr val="accent1"/>
    </a:fillRef>
    <a:effectRef idx="0">
      <a:schemeClr val="accent1"/>
    </a:effectRef>
    <a:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </a:fontRef>
  </p:style>
  <p:stCxn id="${fromShapeId}" idx="${fromConnectionSite}"/>
  <p:endCxn id="${toShapeId}" idx="${toConnectionSite}"/>
</p:cxnSp>`
}

/**
 * 连接点索引说明：
 * 矩形有 4 个连接点：
 * 0 = 上 (top)
 * 1 = 右 (right)
 * 2 = 下 (bottom)
 * 3 = 左 (left)
 */
