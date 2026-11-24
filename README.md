# Word 格式排版助手（Word Format Assistant）

一个面向「本科毕业论文」的 Windows 桌面小工具：**给一份格式混乱的稿子，按学校模板自动重排**；
还能提取模板格式、套封面、画数据图、联网找文献并按 GB/T 7714 排参考文献。

> 它是排版工具，不是 AI 写作工具：所有内容都来自你自己的稿子，程序只改格式。

## 功能

| 页签 | 做什么 |
| --- | --- |
| 排版文档 | 选模板方案 → 一键重排；可选套封面、自动更新目录页码、导出 PDF |
| 新建空白论文 | 按方案直接生成「排好格式的空白论文」，打开即可写 |
| 封面库 | 从模板抽取封面（只留空白表），预览、套用（只改格式 / 重建自动填 / 重建留空） |
| 格式预览 | 不打开 Word 也能看方案长什么样 |
| 图表 | 12 种图（柱/堆叠/折线/饼/环形/散点/箱线/条形/热力/雷达…），读 Excel/CSV，插入文档自动编号 |
| 文献引用 | OpenAlex 检索真实论文 → 句末插上标 [n] → 自动补参考文献；支持 9 种体例与自定义模板；导出 RIS 供 EndNote |
| 提取模板格式 | 从范文提取 22 个部分 / 40+ 样式的方案；提取后自动「重排模板并逐页比对」验收 |

## 下载

| 方式 | 适合谁 | 链接 |
| --- | --- | --- |
| **免安装版 exe（推荐）** | 不想装 Python | [Releases → v1.0.0](https://github.com/sycamore0813/word-/releases/latest) |
| 源码 zip | 想看代码 / 自己改 | 仓库页 Code → Download ZIP |
| git clone | 开发者 | `git clone https://github.com/sycamore0813/word-.git` |

免安装版：下载 zip → 解压到 D 盘或桌面 → 双击 `Word格式排版助手.exe`。
首次运行 Windows 可能提示「已保护你的电脑」，点「更多信息 → 仍要运行」。

## 快速开始

```bat
pip install -r requirements.txt
python app_main.py
```

或直接下载 Releases 里的免安装版（解压双击 `Word格式排版助手.exe`）。

## 命令行

```bat
python -m wordfmt.cli extract 模板.docx --out configs\我的方案.yaml   :: 提取格式（顺带比对）
python -m wordfmt.cli apply 稿件.docx configs\论文格式.yaml           :: 按方案排版
python -m wordfmt.cli newdoc --profile configs\论文格式.yaml            :: 生成排好格式的空白论文
python -m wordfmt.cli skeleton 模板.docx                                :: 只留章节骨架
python -m wordfmt.cli cite 稿件.docx --topic "extrusion cooking"        :: 联网插引用
python -m wordfmt.cli renumber 稿件.docx --check                        :: 检查引用编号
```

## 设计要点

- **样式层重置**：改样式定义 → 段落套样式 → 清直接格式，所以以后在 Word 里改一次样式全文跟着变。
- **提取后自动比对**：用新方案把模板自己重排一遍，逐页像素差异（正常 ≤0.5%），页数必须一致。
- **结构件护栏**：排版前后统计分页符/图片/域/制表符/分节符，任何一类变少就报警。
- **中西文分离**：样式写 `ascii`/`hAnsi`/`eastAsia` + `hint="eastAsia"`，避免全角空格变方框。

## 依赖

`python-docx`、`PyYAML`、`matplotlib`、`numpy`、`openpyxl`、`jieba`、`lxml`、`Pillow`；
可选：`pymupdf`（比对/预览用）、本机 Word（自动更新目录页码、导出 PDF）。

## 说明

- 只支持 Windows + .docx；
- 模板方案在 `configs/*.yaml`，可以手改；
- 文献检索走 OpenAlex / Crossref（无需账号、无需 EndNote）。

## License

MIT
