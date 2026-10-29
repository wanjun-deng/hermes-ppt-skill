# PowerPoint Skill

Hermes Agent 的 `powerpoint-dev` 技能，基于 `python-pptx`。支持用户模板制作、带来源的素材整理、本地模板能力筛选、页计划编译、PPTX 编辑与逐页验收。文件处理可离线运行，视觉模型调用使用 Hermes 已配置的服务。

2026-10-21 起，统一入口为 `SKILL.md`，完整制作流程见
[production-workflow.md](references/production-workflow.md)，视觉报告规则见
[visual-qa.md](references/visual-qa.md)。原有六个 CLI 保持兼容。

新增命令：`pptx_doctor.py`、`pptx_materials.py`、`pptx_templates.py`、
`pptx_workflow.py`。项目 `.venv` 已安装 Word/PDF/Excel 读取依赖；
`uv.lock` 固定了依赖解析，其他机器可用 `uv sync --extra materials` 建立环境。
核心编辑仍可使用现有 `pptx-py`。

## 目录

```text
powerpoint-skill/
├── SKILL.md
├── README.md
├── pyproject.toml
├── schemas/                       # 对外 JSON 契约
├── scripts/
│   ├── pptx_create.py             # powerpoint.create/v1
│   ├── pptx_read.py               # inventory/v2 与图片导出
│   ├── pptx_edit.py               # 通用编辑 CLI
│   ├── pptx_from_template.py      # token 与 layout 追加
│   ├── pptx_from_complete_template.py
│   ├── pptx_render.py             # render/v1
│   ├── pptx_inventory.py          # locator、group 与视觉 inventory
│   ├── pptx_text.py               # 格式保留文本变更
│   └── pptx_common.py             # 原子保存与统一错误
└── tests/
```

模板、提示词、评测运行结果不放在源码目录，统一放到相邻的 `powerpoint-skill-eval/`。

## 关键契约

- 输入：`powerpoint.create/v1`、`powerpoint.template-add/v1`、
  `powerpoint.plan/v1`
- 结构输出：`powerpoint.inventory/v2`、`powerpoint.image-export/v1`、
  `powerpoint.render/v1`
- 变更结果：`powerpoint.create-result/v1`、
  `powerpoint.edit-result/v1`、`powerpoint.template-result/v1`、
  `powerpoint.plan-result/v1`
- 失败输出：`powerpoint.error/v1`

以上 11 个基础 JSON 契约都在 `schemas/` 中有 Draft 2020-12 Schema；
测试不仅检查 Schema 本身，还用真实 CLI 输出做实例校验。

新增 5 个工作流契约：`powerpoint.brief/v1`、`powerpoint.materials/v1`、
`powerpoint.presentation/v1`、`powerpoint.visual-review/v1`、`powerpoint.qa/v1`。
素材记录保留文件 SHA 和位置；制作计划编译到现有 `powerpoint.plan/v1`。
成品验收要求当前 PPTX 与逐页图片哈希一致、全部页面已审阅。
生成但未审阅的文件为 draft；重新生成时保留上一版产物、执行计划与检查记录。

## 基本流程

创建演示文稿（`blank` layout 如果显式提供 `title`，会创建一个普通标题
文本框，不再静默丢弃）：

```json
{
  "schema": "powerpoint.create/v1",
  "slide_size": "16:9",
  "slides": [{"layout": "title", "title": "Annual Review"}]
}
```

```bash
pptx-py scripts/pptx_create.py deck.json out.pptx
```

读取 inventory，并从中复制 template SHA、slide/layout locator 和 shape `recommended_ref`：

```bash
pptx-py scripts/pptx_read.py inspect template.pptx --pretty
pptx-py scripts/pptx_read.py export-images template.pptx --outdir ./images
```

复杂模板 plan 示例：

```json
{
  "schema": "powerpoint.plan/v1",
  "template": {
    "sha256": "COPY_FROM_INVENTORY",
    "slide_count": 19
  },
  "operations": [
    {
      "action": "duplicate",
      "slide": {"slide_id": 256, "index": 0},
      "as": "cover_copy"
    },
    {
      "action": "fill",
      "slide": {"alias": "cover_copy"},
      "text": [{
        "shape": {"placeholder": 0},
        "value": "Annual Review"
      }]
    }
  ]
}
```

```bash
pptx-py scripts/pptx_from_complete_template.py template.pptx plan.json out.pptx
pptx-py scripts/pptx_render.py out.pptx --outdir ./render
```

原始页面使用 inventory 的 `{slide_id,index}` 双重校验。`duplicate`/`append` 可声明 `as`，后续通过 `{"alias":"..."}` 引用，不受删除和重排造成的索引变化影响。Shape 继续使用 inventory 的严格 `recommended_ref`。

## 安全行为

- 创建、编辑和模板输出采用“临时保存 → ZIP/PPTX 重新打开验证 → 原子替换”。
- 未知字段、未知 layout/shape/chart 类型和越界引用不会静默回退。
- 所有可预期失败写入结构化 `powerpoint.error/v1`，并返回非零退出码。
- group 子 shape 会参与文本替换和 locator 解析。
- role inference 默认关闭；只有 plan 显式设置 `allow_inferred_roles: true` 才启用。
- `pptx_render.py` 在隔离目录生成图片，核对页数，并为每页返回 SHA-256 和像素尺寸。

## 开发与验证

环境依赖记录在 `pyproject.toml`。Hermes 环境使用 `pptx-py`：

```bash
PYTHONDONTWRITEBYTECODE=1 pptx-py -m pytest tests/ -q -p no:cacheprovider
```

详细字段、限制和操作建议见 [SKILL.md](SKILL.md)。

首批联网选取的 5 套候选模板位于相邻的
[`powerpoint-template-library/`](../powerpoint-template-library/README.md)，
包含原件、PPTX 主副本、逐页预览、用途标签、严格定位与能力限制。
已填充的图片占位符也会被识别为可导出、可替换的图片；空图片占位符仍需另行插图。
渲染为每次转换使用隔离的 LibreOffice 配置，受限沙箱仍可能需要允许 LibreOffice 运行。

模拟用户测试位于相邻的
[`powerpoint-skill-eval/`](../powerpoint-skill-eval/README.md)。其中首个
case 固化了 `template1.pptx` 的 SHA、19 页、11 个 layout、601 个 slide
shape 和 6 个已知 approximate-bounds warning；可先验证模板 baseline，再对
Agent 生成的 PPTX 执行结构检查与人工逐页视觉评审。
