# 增强模式标签功能设计

## 1. 需求概述

在现有搜索功能基础上，增加"增强模式"开关。开启后，四级行业代码行右侧显示特殊行业分类标签，便于快速识别行业所属的特殊分类体系。

## 2. 数据结构

### 2.1 文件拆分

| 文件 | 说明 |
|------|------|
| `industry_tree_basic.json` | 基础版，保持原有数据结构 |
| `industry_tree_advanced.json` | 增强版，新增 tags 和 tagDetail 字段 |

### 2.2 标签定义（tags 对象）

```javascript
{
  "highTech_mfg": "高技术（制造业）",
  "highTech_svc": "高技术（服务业）",
  "ip密集型": "知识产权密集型产业",
  "strategic": "战略性新兴产业",
  "digital": "数字经济核心产业",
  "pension": "养老产业",
  "culture": "文化产业"
}
```

### 2.3 四级节点数据结构

```json
{
  "name": "化学药品原料药制造",
  "code": "2710",
  "level": 4,
  "description": "◇ 包括对下列...\n◆ 不包括...",
  "tags": ["highTech_mfg"],
  "tagDetail": {}
}
```

- `tags: []` - 空数组表示不属于任何特殊分类
- `tags: ["highTech_mfg", "strategic"]` - 属于多个分类
- `tagDetail: {}` - 空对象表示该行业代码为精确匹配，无需细化

### 2.4 带 * 号的行业处理

当行业代码带 `*`（如 `3919*`）时，表示该类别仅部分活动属于该分类：

```json
{
  "name": "其他计算机制造",
  "code": "3919",
  "level": 4,
  "tags": ["highTech_mfg", "strategic"],
  "tagDetail": {
    "strategic": "仅"其他计算机制造"中的路由器、交换机等网络设备制造属于战略性新兴产业"
  }
}
```

- `tagDetail` 中只存储需要细化的 tag，精确匹配的不存储

## 3. UI 设计（待实现）

- 增强模式按钮置于主题切换按钮旁
- 四级节点右侧显示标签组
- 带 * 号的标签，点击弹出浮窗显示细化说明
- 多标签平行排列

## 4. 数据处理规则

### 4.1 代码匹配规则

1. 优先精确匹配 4 位代码
2. spec.xlsx 中带 `*` 的代码（如 `3919*`）→ 匹配前 4 位（`3919`），并记录 `tagDetail`
3. 行业名称匹配国民经济行业分类标准名称

### 4.2 tagDetail 来源

根据 spec.xlsx 中各 sheet 的"重点说明"或"说明"列提取。

## 5. 实现步骤

1. 解析 spec.xlsx 生成 `industry_tree_advanced.json`
2. 将原 `industry_tree.json` 重命名为 `industry_tree_basic.json`
3. 前端增加增强模式开关
4. 前端渲染标签和浮窗逻辑

## 6. 待确认

- "45432" 具体指什么分类体系？
- UI 标签样式偏好？
