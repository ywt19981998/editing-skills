# 工具说明

Python 一律用共用环境 `~/.venvs/book-review/bin/python`。下面的命令都只读原稿，不修改。

## 符号清单（symbol_inventory.py）——有 PDF 时必跑

```text
# 1. 列清单（纯脚本，几秒）
~/.venvs/book-review/bin/python ~/.claude/skills/formula-proofread/scripts/symbol_inventory.py 书稿.pdf 工作目录/符号清单 [--pages 1-50]
# 2. 模型填好 symbol-table.json 后，自动比对
~/.venvs/book-review/bin/python ~/.claude/skills/formula-proofread/scripts/symbol_inventory.py --apply 工作目录/符号清单/symbol-inventory.json 符号表.json
# 3. 交付前核对每组都有结论
~/.venvs/book-review/bin/python ~/.claude/skills/formula-proofread/scripts/symbol_inventory.py --check 工作目录/符号清单/symbol-inventory.json 工作目录/符号清单/symbol-decisions.json
```

- 读 PDF 文字层（不用 OCR）。还原 Symbol、Euclid Math 等 MathType 字体的私用区编码和 Word 公式的 Unicode 数学字母（𝑣、𝑨 等）。被 PDF 拆成单独一行的上下标，按位置挂回所属的行。
- 列入：正文里的单个字母、公式里的字母、**全部上下标**（i、ij、min、str 等，正体斜体都列）、括注标号 (a)(b)。不列入：正常位置的正体多字母词（DDoS、HGNN 等名称缩写，只计数）、参考文献（从“参考文献”标题到下一章）。
- 重点候选：正文与公式中同一字母字形不一致；公式中同一字母个别处字形与多数不同；正文里的正体单字母；说明性上下标（min、max、str、ref……）用了斜体；MathType 特殊字体（花体、哥特体、空心体）。
- 输出：`symbol-inventory.json`（每组全部出现位置 `occurrences`）、`symbol-inventory.md`、`symbol-table.json`（符号表模板：每个符号的各字形处数和定义句）、`image-regions.json`（图片格式的图，要看图）、`symbol-decisions.json`（判定）。
- `--apply`：按符号表（符号＋可选 position＋scope 适用页 → expected 应有字形）逐处比对，自动判正确／错误／部分错误；没有条目的组标“待判”。已经人工判定的组不会被覆盖。
- `--check`：每组都有结论（正确／错误／部分错误／待复核），错误类都写了理由，否则退出码为 1。
- 实测：《高阶复杂网络系统异常检测》273 页约 5 秒；数学字母约 9 800 处，964 组，符号表 234 个符号；排除参考文献 29 页；图片格式的图 122 处。

## PDF 字形扫描（scan_pdf_glyphs.py）——有 PDF 时必跑

```text
~/.venvs/book-review/bin/python ~/.claude/skills/formula-proofread/scripts/scan_pdf_glyphs.py 书稿.pdf 工作目录/PDF扫描 [--pages 1-50] [--crops]
```

- 从 PDF 的字体信息找候选：微分号 d 斜体（含 dxdy 连写的第二个 d）、自然常数 e 斜体（e 右上方有上标）、e 指数里的虚数单位 i 斜体、函数名斜体（sin、ln、exp、lim、max 等）、数字斜体、π 斜体；并统计 d 的正体、斜体数量，判断全书是否混用。
- **Word 里全是 MathType 公式时，这是唯一能自动查 d、e、i 的手段**（scan_omml.py 读不到 MathType 内容）。
- 输出：`glyph-candidates.json`（每条有页码、bbox、上下文、字体）、`glyph-report.md`、`formula-pages.json`（含数学字体的页，作逐页看图的待审页清单）；`--crops` 为每条候选裁局部图。
- 全部是候选：d 可能是距离、直径变量，i 可能是求和指标，ln 可能是 GB/T 3505 的评定长度，须读定义再判。查不到合成斜体、转曲轮廓字和图片公式，扫描结果不能代替逐页看图。

## Word 公式属性扫描（scan_omml.py）

```text
~/.venvs/book-review/bin/python ~/.claude/skills/formula-proofread/scripts/scan_omml.py 书稿.docx 扫描结果.json
```

- 只读 docx 内部 XML（Word 自带公式 OMML）：函数名是否正体，π、e、d 的正斜体数量，撇号、增量、不等号、省略号、乘号的编码混用，MathType 与 Word 公式混用的位置。JSON 中每类附最多 8 个例子。
- MathType（OLE）公式只统计个数和位置，内容读不到，必须用 PDF 扫描和看图。

## 全文提取（dump_docx_math.py）

`dump_docx_math.py 书稿.docx > 全文.txt`：Word 公式转为可读线性文字（⟦ ⟧ 内为公式），便于数学审查。MathType 内容提取不到。

## PDF 逐字符取证与渲染（prepare_pdf.py）

```text
~/.venvs/book-review/bin/python ~/.claude/skills/formula-proofread/scripts/prepare_pdf.py 书稿.pdf 工作目录/取证 --scale 2
```

保留字符编码、字体、变换矩阵、描边和坐标，渲染整页 PNG；用于判断合成斜体（剪切矩阵）、描边加粗等字体名看不出的情况。不做语义分类，不判错。

## 清单导出（export_review.py）——所有交付必须用它

```text
~/.venvs/book-review/bin/python ~/.claude/skills/formula-proofread/scripts/export_review.py 审校记录.json 交付目录 [--formats pdf|md|both]
```

- 记录格式、版式、正确格式的三种写法见 [清单格式](report.md)。记录有缺字段、编号重复、页码越界、图片缺失、覆盖状态自相矛盾时，不导出并列出问题。
- 原页截图和红框由导出器按每条的 `boxes` 从 `source_pdf` 生成（写法见清单格式第四节）。找不到目标、目标不唯一、字形条目框数少于 `count`、`boxes` 为空时不导出。
- macOS 用宋体（Songti）、Times New Roman，缺字时用 STIX Two Text 补；其他系统用 `--font-dir` 指定字体目录。
- 导出后必须渲染 PDF 目视核对。

## 数学验算（sympy）

共用环境已装 sympy：`import sympy as sp`，常用 `sp.diff`、`sp.simplify`、`sp.factor`、`sp.solve`、`sp.limit`、`sp.series`。设符号时写明假设（如 `sp.symbols('x', positive=True)`），结论中注明。数值代入只能找反例，不能当证明。
