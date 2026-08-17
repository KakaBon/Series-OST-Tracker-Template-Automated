# Series-OST-Tracker-Template-Automated

一个用于记录和整理电视剧 / 分集视频中 OST、配乐及其它音乐出现位置的 Excel 自动化追踪表。

核心使用方式很简单：

- 在 `OST Timeline` 里按剧集和时间点记录每一次音乐出现；
- `OST Collection` 自动整理已经识别的曲目，并统计它们在整部作品和各集中的使用情况；
- 已经收录过的曲目可以通过 Timeline 的 Title 下拉列表再次选择，减少重复填写。

## 适合记录什么

每次音乐出现可以记录：

```text
Episode
Timestamp
Title
Artist
Album
Note
```

例如：

```text
EP 1 | 12:37 | Wayfarer | Ardie Son | Golden Age | 
EP 2 | 08:15 | Wayfarer | Ardie Son | Golden Age |
```

同一首曲在不同位置出现多次，就在 `OST Timeline` 中登记多行；`OST Collection` 会把它们归到同一条曲目记录并进行统计。

## 环境需求

需要：

- Microsoft Excel
- 支持 VBA 宏的桌面环境

工作簿格式为：

```text
.xlsm
```

打开后如果 Excel 显示“宏已被禁用”之类的安全提示，需要允许该工作簿运行宏，否则自动同步、统计、下拉列表、按钮等功能无法正常工作。

## 仓库结构

```text
Series-OST-Tracker-Template-Automated/
├── .gitignore
├── Tool/
│   └── Series-OST-Tracker-Template-Automated.xlsm
├── Example/
│   ├── Payback the Series OST_Identified via AHA, Shazam, YouTube & Google SongSearch.xlsm
│   ├── Payback the Series OST_Identified via AHA, Shazam, YouTube & Google SongSearch - 副本.xlsm
│   ├── Exports/
│   │   └── Payback the Series OST_Identified via AHA, Shazam, YouTube & Google SongSearch_Viewer.html
│   └── OST XLSM CODES/
│       ├── DieseArbeitsmappe.cls
│       ├── modFormatting.bas
│       ├── modOSTEngine.bas
│       ├── Tabelle1.cls
│       └── Tabelle2.cls
├── docs/
|    └── index.html
│       └── Tabelle2.cls
└── README.md
```

### `Tool/`

主工具工作簿：

```text
Series-OST-Tracker-Template-Automated.xlsm
```

普通使用主要操作这个文件。

> 当前这次上传的仓库版本中，`Tool` 里的工作簿已经包含实际 OST 数据，如果要用于新的影视项目，建议先复制一份作为自己的项目文件，再根据需要清理已有业务数据；不要直接把仓库中的原文件当成唯一工作副本。

### `Example/`

以 *Payback the Series* 为实际示例的完整工作簿。

可以用来查看：

- Timeline 应该怎样登记；
- Collection 最终是什么样；
- Total / First Appears / EP 统计如何显示；
- HTML 观看版导出后的效果。

### `Example/Exports/`

已经导出的 HTML 观看版示例。
点击链接查看 HTML 观看版：https://kakabon.github.io/Series-OST-Tracker-Template-Automated/

### `Example/OST XLSM CODES/`

工作簿中的 VBA 代码导出文件。

普通使用工具时不需要操作这个目录。

## 工作簿结构

当前工作簿包含四张表：

```text
OST Timeline
OST Collection
OST Timeline Backup
OST Collection Backup
```

日常主要使用：

```text
OST Timeline
OST Collection
```

工作簿中的一些辅助列会自动隐藏，普通使用时不需要手动维护。

> 当前 VBA 中没有把主表实时自动写入两个 `Backup` 工作表的操作流程，因此不要把这两张表理解为“每次编辑都会自动更新”的实时备份。

## 基本使用流程

### 1. 先复制一份工作簿

不要直接把仓库中的唯一 `.xlsm` 当作长期工作文件。

例如复制并改名：

```text
My Series OST Tracker.xlsm
```

以后在自己的副本中填写。

### 2. 用 Excel 打开并启用宏

打开 `.xlsm` 后允许宏运行。

工作簿打开时会自动整理：

- Timeline 与 Collection 的关联；
- 曲目下拉列表；
- Collection 内容；
- Match / No Match；
- Total；
- First Appears；
- EP 统计。

### 3. 在 `OST Timeline` 登记音乐

主要填写 A–F 列：

```text
A  Episode
B  Timestamp
C  Title
D  Artist
E  Album
F  Note
```

每一次音乐出现使用一行。

例如：

```text
EP 1 | 3:08 | Midnight Mermaids | James Forest | Music for Films |
```

## Episode 怎么写

统计逻辑识别：

```text
EP1
EP 1
EP2
EP 2
...
```

空格不会影响集数识别。

Collection 中的 EP 统计列会根据 Timeline 里出现的最大集数自动延伸。

例如 Timeline 已经出现 `EP 6`，Collection 会建立：

```text
EP1
EP2
EP3
EP4
EP5
EP6
```

## Timestamp 怎么写

可以直接输入已经带冒号的时间：

```text
0:45
12:52
1:03:02
```

也可以快速只输入数字，程序会自动整理。

例如：

```text
45      → 0:45
045     → 0:45
145     → 1:45
1252    → 12:52
10302   → 1:03:02
100302  → 10:03:02
```

如果输入的内容已经有冒号，或不是纯数字，程序不会强行改写。

## 登记一首全新的曲目

如果这首曲还没有被收录：

1. 在 Timeline 新行填写 `Episode` 和 `Timestamp`。
2. 在 `Title` 中输入曲名。
3. 根据需要填写 `Artist`、`Album`、`Note`。
4. 工作簿会根据 Timeline 自动整理 `OST Collection`。
5. 后续统计会随数据变化更新。

一条曲目是否与已有条目完全相同，会综合：

```text
Title
Artist
Album
Note
```

因此同名曲目的不同版本可以通过 Artist / Album / Note 区分。

## 再次登记已经收录的曲目

Timeline 的 `Title` 列带有下拉列表。

当某首曲之后再次出现时：

1. 在新行填写 `Episode` 和 `Timestamp`。
2. 点击 `Title` 单元格的下拉列表。
3. 选择已经收录的曲目。
4. 对应的：

   ```text
   Title
   Artist
   Album
   Note
   ```

   会自动带入。

这样不需要每次重新输入完整曲目信息。

## `OST Collection` 是做什么的

`OST Collection` 是整部作品的曲目汇总表。

主要可看到：

```text
Nr.
Title
Artist
Album
Note
Total
First Appears
EP1
EP2
...
```

### `Total`

这首曲在普通 Timeline 出场记录中的总次数。

### `First Appears`

按 Timeline 中首次记录到的位置显示：

```text
EP - Timestamp
```

例如：

```text
EP 1 - 1:09
```

### `EP1 / EP2 / ...`

显示这首曲在每一集出现了多少次。

## 显示 / 隐藏分集统计

`OST Collection` 中有 EP Stats 按钮。

当分集统计当前显示时，按钮显示：

```text
Hide EP Stats
```

点击后隐藏 `EP1` 到当前最大集数的统计列。

隐藏后按钮会变成：

```text
Show EP Stats
```

再次点击即可展开。

工作簿更新统计时会尽量保持当前展开 / 隐藏状态。

## 修改已经收录的曲目信息

可以直接在 `OST Collection` 中修改：

```text
Title
Artist
Album
Note
```

修改后，与这一条 Collection 记录关联的 Timeline 行会同步更新。

其中：

- `Title` 不能单独清空；
- `Artist`、`Album`、`Note` 可以为空。

如果修改后与另一条曲目的 Title / Artist / Album / Note 完全一致，工作簿会重新整理并归并相同记录。

## 删除 Collection 中的重复 / 版本条目

`OST Collection` 中有：

```text
Delete Entry
```

按钮。

它不是用来随意删除唯一曲目的，而是用于同一个 Title 存在多个 Collection 条目时，将其中一条删除并把相关 Timeline 登记改绑到另一个同名条目。

使用方法：

1. 在 `OST Collection` 中选中准备删除的条目所在行。
2. 点击：

   ```text
   Delete Entry
   ```

3. 如果同一 Title 只有另一个候选条目，Excel 会询问是否把原 Timeline 登记全部替换到那个条目。
4. 如果有多个同名候选，会要求输入要替换到的 `Nr.`。
5. 确认后，原来绑定到被删除条目的 Timeline 行会改为所选条目的 Title / Artist / Album / Note。

如果这个 Title 在 Collection 中只有唯一一条记录，工具会拒绝删除，因为 Timeline 中没有可替换的同名条目。

## 未识别曲目（No Match）

如果某个时间点明显有音乐，但暂时还不知道曲名，可以先登记：

```text
Episode
Timestamp
```

并让：

```text
Title
```

保持为空。

在继续登记后续内容的过程中，工具会为前面尚未识别的记录分配类似：

```text
No-Match-1
No-Match-2
...
```

的临时编号，并写入 `Note`。

`OST Collection` 的 Title 表头会显示：

```text
Match: ...
No Match: ...
```

其中 No Match 按这些临时编号统计。

当以后识别出这条音乐并填写 Title 后，如果 Note 中是工具生成的 No-Match 临时内容，会被清理并重新整理编号。

## Opening Theme / Ending Theme 汇总行

如果 Timeline 的 `Timestamp` 写成：

```text
Opening Theme
Ending Theme
```

这类行会被当作汇总 / 说明行，不计入普通曲目：

```text
Total
EP1 / EP2 / ...
```

的出场次数统计。

适合在表尾单独记录主题曲信息，而不希望它再次增加剧中实际出场次数时使用。

## 导出查看

如果不希望直接分享可编辑的 `.xlsm`，可以使用同组工具：

```text
OST-HTML-Exporter
```

把：

```text
OST Timeline
OST Collection
```

导出成一个浏览器直接打开的 HTML 观看版。
