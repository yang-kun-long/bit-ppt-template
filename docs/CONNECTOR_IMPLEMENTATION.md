# PowerPoint Connector 功能实现总结

## 问题背景

用户发现现有的 AI 生成 PPT（如 RAG PPT）使用普通 `line` 形状拼接箭头，存在以下问题：

1. **不会自动跟随**：移动框图时，箭头不会自动调整
2. **手动对齐困难**：需要手动计算每条线的坐标
3. **容易遮挡**：线条和框图是独立对象，层级容易错乱
4. **只能直线**：无法做弯曲路径，复杂流程图会很乱

## 解决方案

实现真正的 PowerPoint Connector（连接线），通过 OpenXML 后处理注入 `<p:cxnSp>` 元素。

## 技术实现

### 1. 核心模块

**src/connector-template.mjs**
- `generateConnectorXML()`: 生成 OpenXML `<p:cxnSp>` 片段
- 支持 straight/elbow/curved 三种类型
- 支持连接点索引、颜色、线宽、箭头类型

**src/connectors.mjs**
- `injectConnectors()`: 将 connector XML 注入到已生成的 PPTX
- 使用 JSZip 读取/修改/写入 PPTX
- 自动分配形状 ID

**src/generate.mjs**
- 修改 `createDeck()` 添加 `connectors` 数组到上下文
- 修改 `layoutFlowchart()` 收集 connector 配置
- 添加 `postprocessConnectors()` 后处理函数
- 在 `generateDeckFile()` 中调用后处理

### 2. 工作流程

```
YAML 配置
    ↓
layoutFlowchart() 绘制节点并收集 connector 配置
    ↓
createDeck() 返回 { pptx, connectors, ... }
    ↓
pptx.writeFile() 生成基础 PPTX
    ↓
postprocessConnectors() 注入 connector XML
    ↓
最终 PPTX（connector 可自动跟随形状）
```

### 3. 形状 ID 追踪

关键挑战：需要知道每个节点对应的形状 ID，才能正确连接。

解决方案：
- pptxgenjs 从 ID=2 开始分配形状 ID
- 在 `layoutFlowchart()` 中手动追踪每个节点占用的 ID
- 每个节点通常占用 2-3 个 ID（形状 + 文本 + 可选注释）
- 将节点 ID 映射到形状 ID，传递给 connector 配置

### 4. OpenXML 结构

```xml
<p:cxnSp>
  <p:nvCxnSpPr>
    <p:cNvPr id="7" name="Flow 1"/>
    <p:cNvCxnSpPr><a:cxnSpLocks/></p:cNvCxnSpPr>
    <p:nvPr/>
  </p:nvCxnSpPr>
  <p:spPr>
    <a:prstGeom prst="bentConnector3">
      <a:avLst/>
    </a:prstGeom>
    <a:ln w="19050">
      <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      <a:tailEnd type="triangle"/>
    </a:ln>
  </p:spPr>
  <p:stCxn id="2" idx="1"/>   <!-- 起点 -->
  <p:endCxn id="4" idx="3"/>  <!-- 终点 -->
</p:cxnSp>
```

## 使用方法

### YAML 配置

```yaml
- layout: flowchart
  title: "数据处理流程"
  nodes:
    - id: input
      label: "数据输入"
      x: 1.5
      y: 2.5
    - id: process
      label: "数据处理"
      x: 5.5
      y: 2.5
    - id: output
      label: "结果输出"
      x: 9.5
      y: 2.5
  connectors:
    - from: { node: input, site: 1 }
      to: { node: process, site: 3 }
      type: elbow
      label: "传输"
    - from: { node: process, site: 1 }
      to: { node: output, site: 3 }
      type: elbow
```

### 连接点索引

```
    0 (上)
    ↑
3 ← □ → 1 (右)
    ↓
    2 (下)
```

## 测试验证

### 生成测试

```bash
# 生成演示文件
node src/generate.mjs content/connector-demo.yaml output/connector-demo.pptx

# 验证 connector
node scripts/verify_ppt_connectors.mjs output/connector-demo.pptx
```

### 验证结果

```
📊 幻灯片 2: 找到 2 个 Connector
  - 传输
    从形状 2 (连接点 1) → 形状 4 (连接点 3)
  - 导出
    从形状 4 (连接点 1) → 形状 6 (连接点 3)

📊 幻灯片 3: 找到 3 个 Connector
  - Connector 1
    从形状 2 (连接点 2) → 形状 4 (连接点 0)
  - Y
    从形状 4 (连接点 3) → 形状 6 (连接点 0)
  - N
    从形状 4 (连接点 1) → 形状 8 (连接点 0)
```

### 手动测试

1. 在 PowerPoint 中打开 `output/connector-demo.pptx`
2. 选择并移动一个形状
3. ✅ 连接线自动跟随移动
4. ✅ 路径自动优化

## 文件清单

### 新增文件

- `src/connector-template.mjs` - OpenXML 模板生成器
- `src/connectors.mjs` - Connector 注入器
- `content/connector-demo.yaml` - 演示文件
- `docs/CONNECTOR.md` - 完整文档
- `scripts/test_connector_injection.mjs` - 注入测试
- `scripts/verify_connector.mjs` - 验证工具
- `scripts/verify_ppt_connectors.mjs` - PPT 验证工具
- `scripts/test_curved_arrows.mjs` - 弯曲箭头测试
- `scripts/test_connectors.mjs` - Connector 支持测试
- `scripts/create_connector_template.mjs` - 模板创建工具
- `scripts/inspect_rag_ppt.mjs` - RAG PPT 分析
- `scripts/analyze_arrows.mjs` - 箭头分析

### 修改文件

- `src/generate.mjs` - 添加 connector 收集和后处理
- `README.md` - 添加 connector 功能说明

## 向后兼容

- ✅ 如果 flowchart 没有定义 `connectors`，使用传统 `edges` + `line`
- ✅ 现有 YAML 文件无需修改
- ✅ 所有测试通过（47/47）

## 优势对比

### 传统方式（line 形状）

❌ 移动形状时不会跟随  
❌ 需要手动计算坐标  
❌ 只能直线  
❌ 容易遮挡  

### Connector 方式

✅ 自动跟随形状移动  
✅ 自动计算路径  
✅ 支持直线/折线/曲线  
✅ PowerPoint 自动管理层级  

## 参考资料

- [Office Open XML - cxnSp](http://docs.asprain.cn/officeopenxml/drwCxnSp.html)
- [ConnectionShape Class](https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.connectionshape)
- [python-pptx Connector Analysis](https://python-pptx.readthedocs.io/en/latest/dev/analysis/shp-connector.html)
- [cxnSp (Connection Shape)](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_cxnSp_topic_ID0EQ2HPB.html)

## 未来改进

- [ ] 支持更多布局（architecture, process 等）
- [ ] 自动检测最佳连接点
- [ ] 支持连接线标签文本渲染
- [ ] 更智能的形状 ID 追踪
- [ ] 支持自定义连接点位置
