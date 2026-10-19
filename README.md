# 课程 PPT 生成工作流 · Deckflow

用纯文本编写课程，一条命令生成 **PPT＋Word 讲稿＋质量检查报告**。

这是一个内容与引擎分离的课件生产工具，适合持续制作 25–30 分钟的社团专题课。新增一期只需复制模板、填写文本与逐页讲稿、放入素材，**无需编写 Python 排版代码**。内容由你提供，引擎负责排版与校验。

## 可以做什么

- 从 `key<TAB>value` 纯文本生成 16:9 PPT；支持封面、要点、分栏、表格、流程、图解等 17 种布局。
- 内置「白色军事简报」主题：米白纸底、橄榄绿结构、红色强调；主题集中配置，可复制后切换。
- 每页讲稿同时写入 PPT 演讲者备注与 Word。Word 的幻灯片正文**从已经生成的 PPT 反读**。
- 构建自动检查页数、总时长、逐页备注、形状越界、封面像素、讲稿一致性、素材、源码卫生和文字高度。
- 严格指出未知字段、缺失素材、错误页码与不适用的布局字段，避免内容被静默丢弃。
- 自带 17 页、25 分 30 秒的 CQB 示例课，覆盖全部布局，包含三张原创概念图和完整逐页讲稿。

不需要模型 API、API 密钥或联网生成内容。工具不提供 GUI、自动内容创作、人物或场景图像生成、动画或音视频。

## 安装与首次运行

需要 **Python 3.10 或以上**，建议使用虚拟环境：

```bash
git clone https://github.com/blckgum/-PPTworkflow.git
cd ./-PPTworkflow
python -m venv .venv
```

激活环境后安装依赖：

```bash
# macOS / Linux
source .venv/bin/activate
# Windows PowerShell 对应命令：.venv\Scripts\Activate.ps1

pip install -r requirements.txt
python run.py CQB
```

产物位于 `courses/CQB/输出/`：

```text
CQB：空间、压力与信息判断.pptx
CQB：空间、压力与信息判断 讲稿.docx
```

Windows 默认使用 Microsoft YaHei；Linux/macOS 默认使用 Noto Sans CJK SC。请安装相应中文字体，或在 `deckflow/theme.py` 中填写本机实际拥有的字体。只生成 Office 文件不需要 LibreOffice；进行 PDF 渲染验收才需要它和 Poppler。

## 新开一期课

```bash
python run.py --new 城市战
```

仅编辑 `courses/城市战/` 下的以下内容：

1. `course.txt`：标题、机构、目标时长、主题等。
2. `deck.txt`：逐页布局、正文、图片文件名、来源和建议时长。
3. `notes/01.txt` 等：逐页完整讲稿。
4. `素材/`：本课图片；封面至少 2.07 MP，建议至少 4.5 MP。

```bash
python run.py 城市战
```

模板是待填写的骨架，未附封面图片；填入自己的图片并替换占位文字后再构建。详细步骤见 **[从零写一期新课](docs/写一期新课.md)**。

## 命令速查

```bash
python run.py CQB             # 构建 PPT、Word 并校验
python run.py --all           # 构建所有非模板课程
python run.py CQB --check     # 仅检查现有产物
python run.py --all --check   # 检查所有课程
python run.py --new 城市战    # 从空白模板创建，已有目录不会覆盖
```

从其他工作目录调用 `run.py` 也可以；课程位置始终相对于仓库根目录。失败返回非零退出码，输出中文错误信息。含空格的课程名请加引号。

## 文档与结构

| 文档 | 内容 |
|---|---|
| [写一期新课](docs/写一期新课.md) | 从创建课程到交付的完整操作步骤 |
| [内容格式](docs/内容格式.md) | 所有字段、重复规则、格式、默认值与校验 |
| [布局参考](docs/布局参考.md) | 17 种布局的字段、示例与使用建议 |
| [验收记录](docs/验收记录.md) | T1–T4 的命令、结果及验证边界 |
| [原始需求规格](docs/需求规格.md) | 实现依据，保留已确认的决定与技术约束 |

```text
run.py                 一键入口与导入前源码检查
deckflow/              通用引擎、主题、布局与校验
courses/_template/     空白课程骨架
courses/CQB/           完整示例：内容、讲稿与三张图
docs/                  中文使用与维护文档
tests/                 单元测试与隔离的 T1–T4 验收
```

`输出/`、缓存、依赖、PDF 和预览图不会提交到 Git。

## 维护与验证

```bash
python -m unittest discover -s tests -v
python tests/acceptance.py --render
```

完整验收需要 `soffice`、`pdfinfo`、`pdftotext`、`pdftoppm`。Linux 可通过系统包管理器安装 `libreoffice-impress libreoffice-writer poppler-utils fonts-noto-cjk`。工具缺失时渲染验收明确失败；不带 `--render` 只执行程序检查，并明确标记 T4 尚未完整验证。

验收在临时仓库副本里建课和注入故障，不污染实际课程；报告、成品及渲染结果保存在被 Git 忽略的 `.validation/acceptance/`。

文字高度检查使用保守估算，无法替代目标电脑的最终视觉复核。正式交付时仍建议用实际使用的 PowerPoint 打开成品，确认字体可用及显示一致。长内容会自动缩小到主题允许的最小字号；仍放不下则报错，需精简或拆页。

示例遵守严肃题材的表达原则：一个核心观点、每页 3–5 个关键点、图片服务于叙事、注明来源，不伪造战例、人物、数据或引语。三张配图均为原创概念示意，不是历史照片。
