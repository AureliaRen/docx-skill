# Revisions and Comments — XML details

Deep reference for `docx_revisions.py` and `docx_comments.py`. Read this
when you need to reason about the raw WordprocessingML, extend the
scripts, or debug an unusual document. Everyday use only needs SKILL.md.

## 两种机制别混淆

Word 有两套彼此独立的标注系统，很容易混为一谈：

| | 修订（tracked changes） | 批注（comments） |
|---|---|---|
| 存放位置 | `word/document.xml` | `word/comments.xml` |
| 元素 | `<w:ins>`, `<w:del>` | `<w:comment>` + `<w:commentRangeStart/End>` |
| 语义 | 正文**同时包含**两个版本 | 正文一字未改，只是挂了条注记 |
| 接受时 | 择一版本留下 | 删锚点连同注记 |

一份文档可以两者并存，而且会互相牵动：批注可以锚在被删除的区间里，此时接受该删除会把批注变成孤儿——锚点没了，注记还在。

## Tracked changes (w:ins / w:del)

Word records run-level tracked changes as wrapper elements inside a
paragraph (`w:p`), in the `w` namespace
`http://schemas.openxmlformats.org/wordprocessingml/2006/main`:

```xml
<w:p>
  <w:r><w:t>Base </w:t></w:r>
  <w:ins w:id="1" w:author="Editor" w:date="2026-01-02T03:04:05Z">
    <w:r><w:t>inserted text</w:t></w:r>
  </w:ins>
  <w:del w:id="2" w:author="Editor" w:date="2026-01-02T03:04:05Z">
    <w:r><w:delText>deleted text</w:delText></w:r>
  </w:del>
</w:p>
```

Key facts the script relies on:

- Deleted text lives in `w:delText`, not `w:t` — that is why plain text
  extraction naturally shows the "accepted" view (insertions visible,
  deletions hidden).
- Resolution semantics:
  - accept `w:ins` → unwrap (move child runs up, drop the wrapper)
  - reject `w:ins` → remove the wrapper and its contents
  - accept `w:del` → remove the wrapper and its contents
  - reject `w:del` → rename each `w:delText` to `w:t`, then unwrap
- Revisions can appear anywhere block content is allowed: body, table
  cells (nested tables too), headers, footers, text boxes. The script
  iterates the body root plus every header/footer part root with
  `root.iter(W+"ins", W+"del")`, which finds them at any depth.
- `w:id` values are unique per revision *element*, but one logical edit
  session may produce several elements. `accept`/`reject --id` acts on
  exactly the element(s) carrying that id.
- **Revision ids are position-dependent（位置相关），不是稳定标识。**它们由编辑端分配，文档在 Word
  里走一遍往返（打开、保存、或重新编辑）后未必保持原值。要精确接受某几条修订
  的正确姿势是：先 `list` 看清当前 id，紧接着就 `accept`/`reject`，**不要沿用
  上一轮跑出来的 id**——这种失败很安静，改掉的是另一条修订。

  > They are assigned by the editing client and are not stable across a round
  > trip through Word. If you plan to accept specific changes, list first, act
  > immediately, and do not reuse ids from an older run.

Not handled by the script (detected by `docx_read.py --revisions` but
left alone): paragraph-mark revisions (`w:rPr/w:ins` on `w:pPr`), table
row insertions/deletions (`w:trPr/w:ins`), format-change records
(`w:rPrChange`, `w:pPrChange`), and moves (`w:moveFrom`/`w:moveTo`).
Moves are rare from typical editors; if present, treat the file with
Word itself rather than guessing.

## Comments

Three cooperating pieces:

1. **`word/comments.xml`** — one `w:comment` element per comment,
   carrying `w:id`, `w:author`, `w:initials`, `w:date`, and body
   paragraphs. Related from document.xml via the relationship type
   `.../comments` and content type
   `application/vnd...wordprocessingml.comments+xml` (also needs a
   `[Content_Types].xml` override — python-docx's part machinery adds it
   when the part is registered).
2. **Range markers in the story** — `w:commentRangeStart w:id="N"`
   before the anchored runs, `w:commentRangeEnd w:id="N"` after them.
3. **The reference run** — a `w:r` containing `w:commentReference
   w:id="N"`, placed right after the range end; it ties the balloon to
   the location.

`docx_comments.py` behavior:

- **list / delete** always work at the XML level, so they handle files
  from any producer. `anchored_text` is reconstructed by walking each
  part root in document order and collecting `w:t` text between the
  start and end markers for each id.
- **add** first isolates the target text into whole runs. If the match
  starts or ends mid-run, the run is split at the boundary (the split
  copies `w:rPr`, so formatting is preserved). Then:
  - python-docx >= 1.2: the native `document.add_comment(runs, ...)`
    API is used (it creates the comments part, markers, and reference
    run itself).
  - older versions or `--xml`: the script builds `word/comments.xml`,
    registers the part + relationship through the opc layer, and
    inserts the markers/reference manually.
- Deleting a comment removes the `w:comment` element and all three
  marker kinds for that id; the anchored document text is untouched.

锚定由「在文档里找到目标文本、再用 `<w:commentRangeStart>` / `<w:commentRangeEnd>`
加一个 `<w:commentReference>` run 把它包起来」实现。由此有三条后果必须知道：

- 锚定取**第一个**匹配。目标文本出现多次时，批注落在文档序第一处，**不报歧义错**；
  需要精确时请传更长、更具体的串。
- 文本被拆在多个 run 里时（短语中途换格式就会这样，很常见），匹配器必须跨 run 拼接；
  这也正是「加完批注后再编辑、再编辑，锚点会漂」的成因。
- 加批注会**重写整个 document part**，所以批注工作要排在重度重排版之前做完。

Modern Word also writes `commentsExtended.xml` (threading/resolved
state). The scripts neither read nor produce it: replies and "resolved"
flags are invisible here, and comments added by this skill are plain
top-level comments.

## 命令行

修订：

```bash
python scripts/docx_revisions.py list in.docx          # JSON，每条带 w:id
python scripts/docx_revisions.py accept-all in.docx
python scripts/docx_revisions.py reject-all in.docx
python scripts/docx_revisions.py accept in.docx --id 3 # 单条
python scripts/docx_revisions.py reject in.docx --id 7
```

批注：

```bash
python scripts/docx_comments.py list in.docx
python scripts/docx_comments.py add in.docx --target "the phrase to anchor to" --text "note"
python scripts/docx_comments.py add in.docx --target "..." --text "..." --author "Reviewer" --initials RV
python scripts/docx_comments.py delete in.docx --id 1
```

每个命令都接受 `-o/--output`；不给就**原地改**。原地改之前先备份，或者干脆每次都显式给 `-o`。

## 判断文档是否干净

`docx_validate.py` 以 JSON 报问题；一份既无修订也无批注的文档验证结果是
`{"ok": true, "issues": []}`。

交付前必须确认没有残留修订——未接受的修订在收件人眼里就是一堆标记文本（重叠的变更对
尤其难看）。同理，交付前也该确认批注要么已删、要么确实想留。

## 与其他脚本的相互作用

- `docx_edit.py` 编辑的是**当前**文本；对含修订的区域来说，「当前文本」同时含原文与修订版，
  直接编辑会改到不该改的那一份。**先 accept/reject，再编辑。**
- `docx_mathify.py` 就地转 run 并保留修订标记，但**跨 ins/del 边界被拆开的公式不会转**——
  span 探测器看到的是两个碎片，认不成一条公式。同样**先解修订**再做公式。
