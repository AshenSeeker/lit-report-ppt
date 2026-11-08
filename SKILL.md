---
name: lit-report-ppt
display_name: 文献汇报 PPT
display_name_en: Literature Report Deck
description: Turn one journal paper into a Chinese group-meeting literature report deck (.pptx) in the Nanyang Normal University physics template, plus a matching Word speaker script. Use when the user asks for a 文献汇报 / 文献报告 / 组会汇报 / journal club deck from a PDF, or hands over a paper and wants slides plus 讲稿.
description_zh: 把一篇期刊论文做成组会文献汇报 PPT（南阳师院物电学院模板）+ 配套 Word 讲稿。用户提到文献汇报、组会汇报、journal club、把这篇 PDF 做成 PPT 时使用。
description_en: Build a group-meeting literature report deck and speaker script from a journal PDF.
category: writing
version: 1.5.0
---

# 文献汇报 PPT

把一篇 PDF 论文做成能上台念的汇报稿。**两个主产物**：一份 PPT，一份对着每页讲的 Word 讲稿，都从同一个 `content.json` 出。**另有一个可选产物**：答辩预测问答双色 HTML（`qa.json` → `build_qa.py`），见下面「答辩预测问答」。

这套模板量自一份真实的组会汇报稿（28 页），不是凭空设计的。所有几何、字号、颜色都在 `references/style-tokens.md` 里，**别凭感觉改**。

---

## 什么时候用

用户说了这些，就该走这里：
- 「把这篇文献做成汇报 PPT」「组会要汇报这篇」
- 「文献汇报」「journal club」「读书报告」+ 一个 PDF
- 给了论文，问「怎么做成汇报」

---

## 四步走

```
1. 抽 → python scripts/extract_paper.py paper.pdf -o work/
2. 裁 → python scripts/crop_figures.py paper.pdf --crop "3,0.08,0.065,0.93,0.448" -o work/figs/
3. 写 → 填 content.json（内容全在这儿，脚本不替你写内容）
4. 出 → build_deck.py → check_deck.py → build_script.py
```

**别跳步。** 特别是第 3 步 —— 内容是判断活，不是搬运活。

### 1. 抽论文

```bash
python scripts/extract_paper.py paper.pdf -o work/
```

产出 `work/paper.json`：逐页文本、`captions`（图注，已过滤正文引用误抓的）、`abstract`。

PyMuPDF 优先，pypdf 兜底。双栏期刊 pypdf 会丢空格、把 fi/fl 连字拆坏成 `\ufb01`，所以顺序不能反。

先读 `abstract` 和 `captions`，**这是全篇的骨架**。正文通读一遍再动笔。

### 2. 裁图

```bash
# 看这一页有哪些图块，附图注行
python scripts/crop_figures.py paper.pdf --list

# 整页截图（P2 文章首页要这个）
python scripts/crop_figures.py paper.pdf --page1 1 -o work/figs/page01.png
python scripts/crop_figures.py paper.pdf --crop "1,0.035,0.020,0.965,0.430" -o work/figs/page01_top.png --dpi 220

# 裁主图，坐标用 0~1 相对值
python scripts/crop_figures.py paper.pdf --crop "5,0.080,0.065,0.930,0.445" -o work/figs/fig04_pl.png
```

⚠️ **pdf 是位置参数，没有 `--pdf`。**

裁完**必须打开看一眼**：有没有切掉子图、有没有把相邻栏的正文带进来、分辨率够不够（放大到 200% 看字，糊了就 `--dpi 300` 重裁）。

⚠️ 期刊 PDF 的图往往是**跨栏整幅图**，上下还夹着正文。别按「整页宽度」裁，先 `--list` 看图片块的真实边界。

#### 裁图踩过的坑（2026-11-02 一篇稿子上全踩了一遍）

这一节的每一条都是真翻过的车，不是预防性提醒。

1. **图正上方那行红色小节标题不是图。** 期刊（Wiley 系尤其明显）在每页图上面印一行深红斜体小节标题，比如 `Mo-doping in Presence of H₂PO₂`、`Stability & pc-LED Performance`、`Intra-Configurational Spin-Flip (ICSF) Transition`。它的 y 坐标落在「图内容」的范围内，按 bbox 取并集就会把它一起框进来 —— 裁出来的图顶着一条红字残条，放进 PPT 一眼假。**裁完先看顶部 5% 有没有红字。**
2. **红色检测的阈值要严，不然面板编号块会冒充标题。** a) / b) / c) / d) 的编号块也是暗红色。用 `r>110, g<95, b<95` 这种宽松阈值筛「红色像素行」，会把编号块当好几次标题，给出的裁切线全是错的。要 `r>150, g<70, b<70` 再加「整行跨度」条件（连续红像素行数 > 40），或者干脆从 PDF 的 span 里按颜色剔，别在渲染图上猜。
3. **「红字行 + 其后第一个空白行」自动定裁切线不可靠。** 实测在三个页面上给过 `y0_rel = 0.1621 / 0.1575 / 0.1372`，三个数字都落在图内部 —— 照着切会把 a) 面板上半截削掉。**脚本给的线只能当参考起点，最终必须眼看。** 宁可手动多调两次坐标，也别信一条算出来的线。
4. **图的边缘可能就压着轴标签。** 差分图的 x 轴标签 `Decay time (ms)` 贴在图的最底边，裁切线少留 8 个像素就会把字的下半截切掉。底部同样：图的下面紧跟着图注，裁多 5 个像素就把图注的头一行带进来，看起来像「底部有一行碎字」。
5. **一次把所有图都查完，别一张一张试。** 五张图是同一个毛病（顶部红标题 / 底部碎字），一页一页「裁—看—再裁」会把时间拖到没边。正确顺序：全部裁完 → 把每张图的顶部 14% 拼成一张对照图一次看完 → 一次性改完坐标 → 再拼一张复检。
6. **`--crop` 的坐标是 0~1 相对值，别从别的期刊抄数字。** 页面尺寸和栏宽都不一样，抄来的坐标只会把图切坏。先用 `--list` 看图注行的 y，图注上方就是图的下界。

验证用的最小动作（比逐张打开省事得多）：

```python
# 把每张图顶部 14% 拼成一张长图，一眼扫完有没有红标题
from PIL import Image
strips = []
for f in FIGURES:
    im = Image.open(f).convert("RGB")
    w, h = im.size
    strips.append(im.crop((0, 0, w, int(h * 0.14))))
```

底部同理，取最后 10%（`im.crop((0, int(h*0.90), w, h))`）。

### 3. 写 content.json

复制 `assets/content.template.json` 改。字段说明见 `references/page-skeleton.md`。

写内容的三条底线：

1. **图注必须完整**，子图一个不漏，上下角标用 `^{}` / `_{}` 标记。这是规范图第 7 条点名的组内通病。
2. **除第 2 页外，所有英文都要译成中文**（用户 2026-11-02 明确要求）。第 2 页是论文首页截图，全英不译。
3. **数字、单位、化学式一个字都不能想当然**。查不到就不写，别推测。

文案口径见 `references/copywriting.md`。

### 4. 出稿

```bash
python scripts/build_deck.py content.json -o out/deck.pptx
python scripts/check_deck.py out/deck.pptx --content content.json
python scripts/build_script.py content.json -o out/讲稿.docx
```

`check_deck.py` 必须**通过、0 个建议**才算完成。它查六类：出界、重叠、空图注、字体缺失、字号越级、缺图。

⚠️ **渲染验证不能省。** 脚本说没问题不等于看起来没问题。用 LibreOffice Kit 渲出来逐页看：

```bash
node <libreoffice-kit>/lib/cli.js render --input out/deck.pptx --output-dir out/render --pages "1,2,3"
```

`--pages` 落盘文件名从 page-0001 起连号，不保留原页号 —— 渲 `1,5,9` 得到的是 page-0001/0002/0003。

### 汇报人姓名与时间

**默认写 `XX`。** 不要替我编一个人名或日期。

用户提前给了，两条路：写进 `content.json` 的 `meta.presenter` / `meta.date`，或者命令行：

```bash
python scripts/build_deck.py content.json -o out/deck.pptx --presenter "张三" --date "2026年11月2日"
```

`build_script.py` 同样支持这两个参数。

---

## 页序（固定）

```
1  封面              标题 / 汇报人 / 时间
2  文章首页截图       + 作者信息卡      ← 全篇唯一允许全英文
3  摘要精读          ①②③④⑤
4  实验条件/表格
5  合成方法
6+ 主图页 ×N         图 + 完整图注     ← 占全篇 70% 以上
末 结论
末 致谢
```

信息太多就拆两页，**不要缩字号硬塞**。放不下先把图缩小，仍放不下再拆页 —— 顺序别反。

样例（Mo³⁺ 那篇）出来 13 页。主图页占 6 页。

---

## 最容易错的六件事

1. **图注偷懒**。写「如图所示」、只写标题不写子图、图注 40 字打发 —— 全部会被 `check_deck.py` 报。图注是听众唯一能回去查的东西。
2. **把三件事塞一页**。宁可多开一页。用户原话：「你太强迫症了，如果不够放就再多开一页呗」。图注放不下会自动续页，**续页上必须有完整的「图 + 注」**，不能只剩注。
3. **用 Unicode 上下标**。`Cr³⁺` 换字体就散，复制到论文里变乱码。必须写 `Cr^{3+}` 让脚本转真上下标。
4. **加了样本里没有的东西**。提炼条、橙色 callout、页码、手画下划线 —— 原稿没有的全不加。用户说过「样本外元素全删，严格照样本」。要加先问。
5. **裁图没看图就交**。顶部红标题残条、底部图注碎字、轴标签被切半个字 —— 这三种都躲得过 `check_deck.py`（它只看形状，不看图里有什么），只能靠眼睛。详见上面「裁图踩过的坑」。
6. **默认校徽校名给了别人**。封面顶栏是**南阳师范学院物理与电子工程学院**，校徽 `assets/logo.png` 和两行校名（`scripts/build_deck.py:53` 与 `:58`）都是硬编码。用的人不是这个学校时**三处都得换**，只换 logo 会让封面顶着别人的校名，很尴尬。开做之前先问一句「你是哪个学校/学院的」；用户不打算换就别擅自改。

---

## 讲稿

`build_script.py` 出的 Word 讲稿，形态是**逐页讲解**：一页 PPT 一段话，一张图两三句，学术严肃但不堆术语。

**不要**写「开场白 / 核心贡献 / 必讲概念 / 可能被问的问题」那种备忘记 —— 用户明确说过只要对 PPT 内容和图片做讲解（2026-11-02）。

⚠️ **讲稿不必和 PPT 逐页一一对齐**（用户 2026-11-02）。两页图讲同一件事，合一段讲更顺。想精确指定就写 `"p": "7-8"`，Word 里会印 `P7–P8`。

写法规范见 `references/script-guide.md`。

**不写进 PPT 备注栏**（用户决定）。备注栏放映时看不见，念的时候还得另开窗口，不如单独一份 Word 放旁边。

---

## 答辩预测问答（可选第三产物）

除了 PPT 和讲稿，还能再出一份**答辩预测问题与答案**，形态是一份双色 HTML：问题蓝、答案绿，表格整齐。用户要「双色版」时走这里。

**它和前面两个产物是同一套论文材料**，只是多了一份 `qa.json` 和一步渲染：

```bash
# 1. 写 qa.json（复制 assets/qa.template.json 改）
#    ↑ 和 content.json 一样，这是判断活，脚本不替你写内容
# 2. 出稿
python scripts/build_qa.py qa.json -o out/答辩预测问题.html
```

`build_qa.py` **只管排版**，CSS 是照着一份真实成品逐字抄下来的（`--q:#1d4ed8` / `--a:#15803d`、`.wrap max-width:1020px`、`col.q 43%`、斑马纹），**别凭感觉改颜色或间距** —— 改一处就和 `examples/答辩预测问题.html` 对不上。`qa.json` 里写 `^{3+}` / `_{2}`，脚本转成真的 `<sup>` / `<sub>`。

### 五条硬要求（照抄，别自己发挥）

1. **只要「问题 + 直接答案」两栏。** 不要「考察点」「为什么问」「应答解析」「参考思路」等任何附加内容 —— 用户明确点名不要。
2. **答案 ≤30 字**（化学式、符号、字母不计）。言简意赅，直接给结论，不铺陈。`build_qa.py` 会数汉字数，超了给警告但不阻断，自己回头删。
3. **问题与答案不同颜色** —— 这正是用 HTML 而不是纯 txt 的原因。颜色脚本固定，不用你调。
4. **排版整齐**：表格、对齐、上下角标用 `^{}` / `_{}`（别直接粘 Unicode 上下标，换环境会散）。
5. **提问方向必须覆盖这三类**：① 为什么选这个体系（基质、掺杂离子、合成路线为什么这么定）② 国内外这个方向哪里做得多（点名课题组/地区，别泛泛说「研究很多」）③ 能否换成别的材料、替换后性质怎么变（引申出对比）。

### 内容从哪来

通读 `paper.json` 的正文和图注，尤其是：合成步骤里不能省的那一步、表格里投料与实测对不上的数、每张主图反常的地方（能级线为什么水平、峰位为什么不随温度动、XRD 重合为什么还要做 XAS、效率与功率是不是同一电流测的）。**数字、条件、单位一个字都不能想当然**，和写 `content.json` 同一条底线。

⚠️ **只出 HTML，不另出纯文本 txt**（用户要的「就最后那个双色版」）。样例见 `examples/答辩预测问题.html`（双击即可在浏览器预览）、`examples/答辩预测问题.png`（截图）、`examples/qa.json`（可改的源数据）。

---

## 文件

```
SKILL.md                        本文件
references/style-tokens.md      几何/字号/颜色/图注容量（硬参数，改前先问）
references/page-skeleton.md     七种页型的字段结构与写法
references/script-guide.md      讲稿写作规范
references/copywriting.md       文案口径：什么话能说、什么话不能说
assets/template.pptx            空模板（母版+13个版式，无页面）
assets/logo.png                 校徽（默认南阳师范学院，自用须换）
assets/content.template.json    content.json 全字段样板
assets/qa.template.json          qa.json 字段样板（答辩预测问答）
examples/                       成品示例：PPT / 讲稿截图 + 源文件 + 双色问答 HTML 与预览图
package.json / cordis.patch.yml / lib/index.js   DSH 插件市场用：宿主侧注册本 skill，不碰出稿脚本
scripts/lr_common.py            设计常量 + 排版原语（改参数改这里）
scripts/extract_paper.py        抽文本与图注
scripts/crop_figures.py         裁图
scripts/build_deck.py           生成 PPT
scripts/check_deck.py           自检
scripts/build_script.py         生成 Word 讲稿
scripts/build_qa.py             生成双色答辩预测问答 HTML
```

---

## 环境

依赖：`python-pptx`、`python-docx`、`Pillow`、`pymupdf`（优先）、`pypdf`（兜底）。

```bash
pip install -r requirements.txt
```

字体用**本机有的**：Times New Roman、楷体、等线、Verdana。原稿的「华文楷体」本机没有，已用「楷体」顶替 —— 别去装字体，`check_deck.py` 会照实际可用字体检查。

⚠️ Windows 下 Python 默认 stdout 是 gbk，打印中文会炸 `UnicodeEncodeError`。所有脚本顶部已有 `sys.stdout.reconfigure(encoding="utf-8")`，自己写新脚本记得加。