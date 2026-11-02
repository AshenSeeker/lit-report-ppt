# 文献汇报 PPT

把一篇期刊论文做成能上台念的中文组会汇报稿。**两个产物**：一份 `.pptx`，一份对着每页讲的 Word 讲稿 —— 都从同一个 `content.json` 出，所以两边讲的不会打架。

模板量自一份真实的组会汇报稿（28 页），所有几何、字号、颜色都是逐形状量出来的实测值，不是凭感觉设的。

---

## 它解决什么问题

组会文献汇报有固定的套路：封面 → 论文首页截图 → 摘要精读 → 实验条件 → 合成方法 → 一张一张主图配图注 → 结论 → 致谢。每次都要重做一遍，最容易翻车的是**图注偷懒**（写「如图所示」、漏子图、上下角标打不出来）—— 这恰恰是听的人唯一能回去查的东西。

这个 skill 把版式固定下来，把容易漏的检查写成脚本：

- 图注少于 40 字、缺子图交代 → 自检时报错
- 图注放不下 → 自动开续页，**续页上仍然带同一张图**（只挪注不挪图，讲这页时图在旁边根本看不到）
- 化学式写 `Cr^{3+}`，脚本转成真上下标（写 Unicode 的 `Cr³⁺` 换个字体就散）
- 出界、重叠、字体缺失、字号越级 → 逐页扫描

---

## 安装

```bash
pip install -r requirements.txt
```

用作 [DSH](https://github.com/) skill 时，把整个目录放进 `~/.dsh/skills/lit-report-ppt/`，DSH 会自动注册。也可以脱离 DSH 直接跑脚本 —— 模板与校徽走 `__file__` 解析，跟当前目录无关。

---

## 用

### 四步走

```bash
# 1. 抽文本和图注
python scripts/extract_paper.py paper.pdf -o work/

# 2. 裁图（先 --list 看图片块的真实边界）
python scripts/crop_figures.py paper.pdf --list
python scripts/crop_figures.py paper.pdf --crop "5,0.080,0.065,0.930,0.445" -o work/figs/fig04.png

# 3. 填 content.json（复制 assets/content.template.json 改）

# 4. 出稿
python scripts/build_deck.py content.json -o out/deck.pptx
python scripts/check_deck.py out/deck.pptx --content content.json
python scripts/build_script.py content.json -o out/讲稿.docx
```

`check_deck.py` 必须**通过、0 个建议**才算完成。

### 汇报人姓名与时间

默认写 `XX`。提前知道了就：

```bash
python scripts/build_deck.py content.json -o out/deck.pptx --presenter "张三" --date "2026年11月2日"
```

或者写进 `content.json` 的 `meta.presenter` / `meta.date`。

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

---

## 主图页有两种排法

不是风格选择，是**按图注长度实际算行数**自动选：

| 排法 | 图区 | 图注栏 | 适合 |
|---|---|---|---|
| A 底栏 | 11.71 × 6.20 | 底栏 11.80 × 0.92 | 短图注，图满宽 |
| B 右栏 | 8.45 × 7.16 | 右栏 2.75 × 7.14 | 长图注，图靠左 |

两栏都放不下 → 开续页。

---

## 自检查什么

`check_deck.py` 扫六类：

1. **出界** —— 形状跑出画布，多半被切掉了
2. **重叠** —— 两个形状压在一起（阈值 0.15in）
3. **空图注** —— 少于 40 字
4. **字体缺失** —— 对照 `C:\Windows\Fonts` 实际可用字体
5. **字号越级** —— 不在白名单 `{66,60,40,22,20,18,16,14,13,11,10}` 里
6. **缺图** —— `image` 指向的文件不存在

---

## 目录

```
SKILL.md                        主入口：什么时候用、四步走、最容易错的四件事
requirements.txt                python 依赖
references/style-tokens.md      几何/字号/颜色/图注容量（硬参数，改前先问）
references/page-skeleton.md     七种页型的字段结构与写法
references/script-guide.md      讲稿写作规范
references/copywriting.md       文案口径：什么话能说、什么话不能说
assets/template.pptx            空模板（母版 + 10 个版式，无页面）
assets/logo.png                 校徽
assets/content.template.json    content.json 全字段样板
scripts/lr_common.py            设计常量 + 排版原语（改参数改这里）
scripts/extract_paper.py        抽文本与图注
scripts/crop_figures.py         裁图
scripts/build_deck.py           生成 PPT
scripts/check_deck.py           自检
scripts/build_script.py         生成 Word 讲稿
```

---

## 关于字体

模板原稿用的是「华文楷体」，多数 Windows 机器没装。本仓库统一用**楷体**（`simkai.ttf`，Windows 自带）顶替，`check_deck.py` 按实际可用字体检查。

另外两个字体是系统自带的，不用管：Times New Roman（西文正文）、Verdana（校名）、等线（序号）。

---

## 已知取舍

- **讲稿不写进 PPT 备注栏。** 备注栏放映时看不见，念的时候还得另开窗口，不如单独一份 Word 放旁边。
- **讲稿不必和 PPT 逐页一一对齐。** 两页图讲同一件事，合一段讲更顺。用 `"p": "7-8"` 指定。
- **左上角竖排标签只留字、不画框。** 外框已删，文字保留并改为楷体 11pt 正文色，竖排方向不变（只出现在 B 排法页，A 排法左边塞不下）。`content.json` 里的 `tag` 字段同时给讲稿生成器当页标题兜底。
- **不主动加原稿没有的元素。** 提炼条、橙色 callout、页码都拿掉了。要加先问。

---

## License

MIT
