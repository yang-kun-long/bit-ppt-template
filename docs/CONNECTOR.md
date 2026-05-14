# PowerPoint Connector 支持

## 概述

bit-ppt-template 现在支持真正的 PowerPoint Connector（连接线），可以自动跟随形状移动。这是通过后处理注入 OpenXML `<p:cxnSp>` 元素实现的。

## 功能特性

### ✅ 已实现

- **自动跟随**: Connector 会自动跟随连接的形状移动
- **多种类型**: 支持直线 (straight)、折线 (elbow)、曲线 (curved)
- **连接点控制**: 可以指定形状的哪个连接点（上/右/下/左）
- **样式定制**: 支持颜色、线宽、箭头类型
- **YAML 配置**: 在 flowchart 布局中通过 `connectors` 字段定义

### ❌ 传统方式的问题

使用普通 `line` 形状的问题：
- 移动形状时，连接线不会跟随
- 需要手动计算坐标
- 只能绘制直线，无法自动折线

## 使用方法

### 在 YAML 中定义 Connector

```yaml
slides:
  - layout: flowchart
    title: "数据处理流程"
    nodes:
      - id: input
        label: "数据输入"
        x: 1.5
        y: 2.5
        color: "006C39"

      - id: process
        label: "数据处理"
        x: 5.5
        y: 2.5
        color: "0066CC"

      - id: output
        label: "结果输出"
        x: 9.5
        y: 2.5
        color: "A13F3D"

    connectors:
      - from: { node: input, site: 1 }    # 从 input 的右侧 (site: 1)
        to: { node: process, site: 3 }    # 到 process 的左侧 (site: 3)
        type: elbow                        # 折线类型
        label: "传输"                      # 连接线名称

      - from: { node: process, site: 1 }
        to: { node: output, site: 3 }
        type: elbow
        label: "导出"
```

### 连接点索引

矩形形状有 4 个连接点：

```
        0 (上)
        ↑
    3 ← □ → 1 (右)
        ↓
        2 (下)
```

- `0` = 上 (top)
- `1` = 右 (right)
- `2` = 下 (bottom)
- `3` = 左 (left)

### Connector 类型

- `straight`: 直线
- `elbow`: 折线（一个拐角）
- `curved`: 曲线

### 完整配置选项

```yaml
connectors:
  - from:
      node: source_id      # 起点节点 ID
      site: 1              # 连接点索引 (0-3)
    to:
      node: target_id      # 终点节点 ID
      site: 3              # 连接点索引 (0-3)
    type: elbow            # 连接线类型
    color: "000000"        # 线条颜色 (RGB hex)
    width: 19050           # 线条宽度 (EMU, 19050 = 0.75pt)
    arrowType: triangle    # 箭头类型: none | triangle | arrow
    label: "连接线名称"     # 可选的名称/标签
```

## 技术实现

### 架构

```
YAML 配置
    ↓
layoutFlowchart() 收集 connector 配置
    ↓
createDeck() 返回 connectors 数组
    ↓
generateDeckFile() 调用 postprocessConnectors()
    ↓
injectConnectors() 注入 OpenXML
    ↓
最终 PPTX 文件
```

### 核心模块

1. **src/connector-template.mjs**
   - `generateConnectorXML()`: 生成 OpenXML 片段

2. **src/connectors.mjs**
   - `injectConnectors()`: 注入 connector 到 PPTX
   - `extractConnectorsFromSlide()`: 从 YAML 提取配置

3. **src/generate.mjs**
   - `layoutFlowchart()`: 收集 connector 配置
   - `postprocessConnectors()`: 后处理入口

### OpenXML 结构

```xml
<p:cxnSp>
  <p:nvCxnSpPr>
    <p:cNvPr id="7" name="Flow 1"/>
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
    <a:prstGeom prst="bentConnector3">
      <a:avLst/>
    </a:prstGeom>
    <a:ln w="19050">
      <a:solidFill>
        <a:srgbClr val="000000"/>
      </a:solidFill>
      <a:headEnd type="none"/>
      <a:tailEnd type="triangle"/>
    </a:ln>
  </p:spPr>
  <p:stCxn id="2" idx="1"/>   <!-- 起点: 形状 2, 连接点 1 -->
  <p:endCxn id="4" idx="3"/>  <!-- 终点: 形状 4, 连接点 3 -->
</p:cxnSp>
```

## 示例

### 简单流程

```yaml
- layout: flowchart
  title: "三步流程"
  nodes:
    - id: step1
      label: "步骤 1"
      x: 2
      y: 3
    - id: step2
      label: "步骤 2"
      x: 6
      y: 3
    - id: step3
      label: "步骤 3"
      x: 10
      y: 3
  connectors:
    - from: { node: step1, site: 1 }
      to: { node: step2, site: 3 }
      type: straight
    - from: { node: step2, site: 1 }
      to: { node: step3, site: 3 }
      type: straight
```

### 分支流程

```yaml
- layout: flowchart
  title: "条件分支"
  nodes:
    - id: start
      label: "开始"
      x: 6
      y: 1.5
    - id: decision
      label: "判断"
      x: 6
      y: 3.5
    - id: yes
      label: "是"
      x: 3
      y: 5.5
    - id: no
      label: "否"
      x: 9
      y: 5.5
  connectors:
    - from: { node: start, site: 2 }
      to: { node: decision, site: 0 }
      type: straight
    - from: { node: decision, site: 3 }
      to: { node: yes, site: 0 }
      type: elbow
      label: "Y"
    - from: { node: decision, site: 1 }
      to: { node: no, site: 0 }
      type: elbow
      label: "N"
```

## 测试

### 生成测试文件

```bash
# 生成演示文件
node src/generate.mjs content/connector-demo.yaml output/connector-demo.pptx

# 验证 connector
node scripts/verify_ppt_connectors.mjs output/connector-demo.pptx
```

### 手动测试

1. 在 PowerPoint 中打开生成的文件
2. 选择并移动一个形状
3. 观察连接线是否自动跟随移动
4. 尝试调整连接线的路径（PowerPoint 会自动优化）

## 兼容性

### 向后兼容

- 如果 flowchart 布局**没有**定义 `connectors` 字段，会使用传统的 `edges` 和 `line` 形状
- 现有的 YAML 文件无需修改即可继续工作

### PowerPoint 版本

- ✅ PowerPoint 2016 及以上
- ✅ PowerPoint for Mac
- ✅ PowerPoint Online
- ⚠️ 旧版本可能不支持某些 connector 类型

## 限制

1. **形状 ID 追踪**: 当前实现假设每个节点占用固定数量的形状 ID，如果布局逻辑改变可能需要调整
2. **仅支持 flowchart**: 目前只在 flowchart 布局中实现，其他布局仍使用传统方式
3. **连接点数量**: 不同形状的连接点数量不同，需要查阅 PowerPoint 文档

## 未来改进

- [ ] 支持更多布局类型（architecture, process 等）
- [ ] 自动检测最佳连接点
- [ ] 支持自定义连接点位置
- [ ] 支持连接线标签文本
- [ ] 更智能的形状 ID 追踪机制

## 参考资料

- [Office Open XML - cxnSp](http://docs.asprain.cn/officeopenxml/drwCxnSp.html)
- [ConnectionShape Class](https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.connectionshape)
- [python-pptx Connector Analysis](https://python-pptx.readthedocs.io/en/latest/dev/analysis/shp-connector.html)
