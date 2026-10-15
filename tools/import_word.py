#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
《消失的灵气》Word → Markdown 反向导入工具（tools/export_word.py 的逆操作）

在 Word 里改完稿后，用这个脚本把改动写回 novel-project/ 的 md 源文件。
切分依据：每个源文件的首行 `# 标题` 在导出时成为固定层级的标题（分类文档=Heading 2，
全量合集=Heading 3），且 95 个标题全局唯一，因此可无损切回。

用法：
    python3 tools/import_word.py                      # 干跑：解析 + 与现有 md 逐字比对，不落盘
    python3 tools/import_word.py --stage              # 把还原出的 md 写到 Word回写/ 暂存目录
    python3 tools/import_word.py --apply              # 直接写回 novel-project/（先提交！）
    python3 tools/import_word.py --only 正文 设定      # 只处理部分文档
    python3 tools/import_word.py -f 某份.docx          # 处理指定 docx

默认处理 Word导出/ 下的 4 个分类文档；全量合集需显式指定，避免与分类文档重复写回。
"""
import argparse
import re
import sys
from collections import Counter
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
BASE = ROOT / 'novel-project'
SRC_DIR = ROOT / 'Word导出'
STAGE_DIR = ROOT / 'Word回写'
BOOK = '消失的灵气'
FOLDERS = ['正文', '大纲', '设定', '状态']
HEADING_RE = re.compile(r'^Heading ([1-6])$')
LIST_MARK = re.compile(r'^(?:[•·]\s{1,2}|(\d+)\.\s)')
CODE_FONT = 'Courier New'


# ===================== 源文件索引 =====================
def top_heading(path):
    for line in path.read_text(encoding='utf-8').splitlines():
        m = re.match(r'^#\s+(.*?)\s*$', line)
        if m:
            return m.group(1)
    return path.stem


def build_index():
    """{首行标题: 相对路径}"""
    idx = {}
    for folder in FOLDERS:
        for p in sorted((BASE / folder).rglob('*.md')):
            title = top_heading(p)
            if title in idx:
                raise SystemExit('标题不唯一，无法反向切分：%s 与 %s' % (idx[title], p))
            idx[title] = str(p.relative_to(BASE))
    return idx


# ===================== docx 读取 =====================
def run_fmt(run):
    return {'bold': bool(run.bold), 'italic': bool(run.italic),
            'strike': bool(run.font.strike),
            'code': run.font.name == CODE_FONT}


def para_runs(p):
    """返回 [(文本, 格式)]，已合并相邻同格式 run；'\\n' 表示 Word 硬换行。"""
    out = []
    for r in p.runs:
        text = r.text or ''
        if not text:
            continue
        fmt = run_fmt(r)
        if out and out[-1][1] == fmt:
            out[-1][0] += text
        else:
            out.append([text, fmt])
    return [(t, f) for t, f in out]


def is_rule(p):
    ppr = p._p.find(qn('w:pPr'))
    return ppr is not None and ppr.find(qn('w:pBdr')) is not None


def body_blocks(doc):
    """按文档顺序产出 ('p', Paragraph) / ('table', Table)。"""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield ('p', Paragraph(child, doc))
        elif child.tag == qn('w:tbl'):
            yield ('table', Table(child, doc))


# ===================== Markdown 生成 =====================
def inline_md(runs, in_table=False):
    parts = []
    for text, f in runs:
        if not text:
            continue
        t = text
        if in_table:
            t = t.replace('|', '\\|')
        if f['code']:
            t = '`' + t + '`'
        else:
            t = t.replace('`', '\\`') if '`' in t else t
        if f['bold']:
            t = '**' + t + '**'
        if f['italic']:
            t = '*' + t + '*'
        if f['strike']:
            t = '~~' + t + '~~'
        parts.append(t)
    return ''.join(parts)


def guard_line_start(line):
    """避免还原出的行首被误认为块语法。"""
    if re.match(r'^(#|>|```|~~~|\*\s|-\s|\+\s|\d+\.\s)', line):
        return '\\ ' + line[0] + line[1:] if line[0] in '#>-' else line
    return line


class Restorer:
    def __init__(self, index, file_level, prose, folder=None):
        self.index = index
        self.file_level = file_level
        self.prose = prose
        self.folder = folder
        self.files = {}          # relpath -> [markdown lines]
        self.new_files = {}      # relpath -> [markdown lines]
        self.new_titles = {}     # relpath -> Word 里的标题
        self.cur = None
        self.target = None
        self.stats = Counter()

    # ---- 输出目标 ----
    def new_path(self, title):
        m = re.match(r'^第\s*(\d+)\s*章[\s　]*(.*)$', title)
        if m:
            stem = '第%03d章-%s' % (int(m.group(1)), m.group(2).strip() or '未命名')
        else:
            stem = re.sub(r'[\\/:*?"<>|]', '', title).strip() or '未命名'
        folder = self.folder or '__new__'
        rel = '%s/%s.md' % (folder, stem)
        n = 2
        while (BASE / rel).exists() or rel in self.new_files:
            rel = '%s/%s(%d).md' % (folder, stem, n)
            n += 1
        return rel

    def start_file(self, title):
        rel = self.index.get(title)
        if rel:
            self.cur = rel
            self.target = self.files.setdefault(rel, [])
        else:
            self.cur = self.new_path(title)
            self.target = self.new_files.setdefault(self.cur, [])
            self.new_titles[self.cur] = title
            self.stats['新文件'] += 1
        self.stats['文件'] += 1
        self.emit('# ' + title)

    def emit(self, line=''):
        if self.target is not None:
            self.target.append(line)

    # ---- 块级 ----
    def paragraph(self, p):
        style = p.style.name if p.style else 'Normal'
        m = HEADING_RE.match(style)
        if m:
            lvl = int(m.group(1))
            text = inline_md(para_runs(p)).strip()
            if not text:
                return
            if lvl < self.file_level:
                self.cur = self.target = None   # 分组/板块标题，不属于任何文件
                return
            if lvl == self.file_level:
                self.start_file(text)
                return
            self.emit()
            self.emit('#' * (lvl - self.file_level + 1) + ' ' + text)
            return
        if self.target is None:
            return
        runs = para_runs(p)
        text = inline_md(runs)
        if is_rule(p):
            if text.strip():
                self.emit(text)
            self.emit('---')
            self.emit()
            return
        if not text.strip():
            return
        if style in ('Quote', 'Intense Quote'):
            for ln in text.split('\n'):
                self.emit('> ' + guard_line_start(ln.strip()))
            self.emit()
            return
        depth = self.list_depth(p)
        mm = LIST_MARK.match(text)
        if mm:
            body = text[mm.end():]
            mark = '%d. ' % int(mm.group(1)) if mm.group(1) else '- '
            self.emit('  ' * depth + mark + guard_line_start(body.strip()))
            return
        # 普通段落：Word 硬换行还原为同段内的续行
        lines = [ln.strip() for ln in text.split('\n')]
        for i, ln in enumerate(lines):
            if not ln:
                continue
            self.emit(('  ' * depth if depth else '') +
                      ('' if i == 0 else '  ') + guard_line_start(ln))
        self.emit()

    @staticmethod
    def list_depth(p):
        try:
            left = p.paragraph_format.left_indent
            first = p.paragraph_format.first_line_indent
            if left is None or first is None:
                return 0
            from docx.shared import Cm
            step = int(Cm(0.7))
            base = int(Cm(0.75))
            return max(0, int(round((int(left) - base) / step)))
        except Exception:
            return 0

    def table(self, t):
        if self.target is None:
            return
        rows = []
        for r in t.rows:
            cells = []
            for c in r.cells:
                chunks = []
                for p in c.paragraphs:
                    runs = para_runs(p)
                    if runs:
                        chunks.append(inline_md(runs, in_table=True))
                cells.append('<br>'.join(x for x in chunks if x.strip()).strip() or ' ')
            rows.append(cells)
        if not rows:
            return
        width = max(len(r) for r in rows)
        for r in rows:
            r += [' '] * (width - len(r))
        self.emit('| ' + ' | '.join(rows[0]) + ' |')
        self.emit('|' + '---|' * width)
        for r in rows[1:]:
            self.emit('| ' + ' | '.join(r) + ' |')
        self.emit()

    def codeblock(self, ps):
        if self.target is None:
            return
        self.emit('```')
        for p in ps:
            self.emit(''.join(t for t, _ in para_runs(p)).rstrip('\n'))
        self.emit('```')
        self.emit()

    def feed(self, doc):
        blocks = list(body_blocks(doc))
        i = 0
        while i < len(blocks):
            kind, obj = blocks[i]
            if kind == 'table':
                self.table(obj)
                i += 1
                continue
            p = obj
            runs = para_runs(p)
            only_code = bool(runs) and all(f['code'] for _, f in runs) and \
                (p.style.name if p.style else '') == 'Normal'
            if only_code:
                j = i
                chunk = []
                while j < len(blocks) and blocks[j][0] == 'p':
                    q = blocks[j][1]
                    rr = para_runs(q)
                    if rr and all(f['code'] for _, f in rr) and \
                            (q.style.name if q.style else '') == 'Normal':
                        chunk.append(q)
                        j += 1
                    else:
                        break
                self.codeblock(chunk)
                i = j
                continue
            self.paragraph(p)
            i += 1


def kind_of(line):
    if re.match(r'^#{1,6} ', line):
        return 'h'
    if line == '---':
        return 'hr'
    if line.startswith('|'):
        return 'tbl'
    if line.startswith('```'):
        return 'fence'
    if line.startswith('> '):
        return 'quote'
    if LIST_MARK.match(line):
        return 'li'
    if line.startswith('  ') and line.strip():
        return 'cont'
    return 'p'


GLUE = {('li', 'li'), ('li', 'cont'), ('cont', 'cont'), ('tbl', 'tbl'), ('quote', 'quote')}


def tidy(lines):
    """块与块之间补且只补一个空行；列表项与其续行、表格行之间保持紧邻；代码围栏内不改动。"""
    res, in_fence = [], False
    for ln in lines:
        stripped = ln.rstrip()
        if stripped == '```':
            in_fence = not in_fence
            res.append(stripped)
            continue
        if in_fence:
            res.append(stripped)
            continue
        if not stripped.strip():
            if res and res[-1].strip():
                res.append('')
            continue
        if res and res[-1].strip():
            pair = (kind_of(res[-1]), kind_of(stripped))
            if pair not in GLUE:
                res.append('')
        res.append(stripped)
    while res and not res[-1].strip():
        res.pop()
    return res


# ===================== 保真度比对 =====================
HAN = re.compile(r'[一-鿿]')


def han(s):
    return ''.join(HAN.findall(s))


def compare(index, restored):
    """还原出的 md 与磁盘上的源文件逐行比对汉字序列。"""
    problems = []
    for rel, lines in sorted(restored.items()):
        src = (BASE / rel).read_text(encoding='utf-8')
        doc = han('\n'.join(tidy(lines)))
        miss = []
        for line in src.splitlines():
            key = han(line)
            if len(key) >= 4 and key not in doc:
                miss.append(line)
        extra = []
        src_han = han(src)
        for line in lines:
            key = han(line)
            if len(key) >= 4 and key not in src_han:
                extra.append(line)
        ratio = len(doc) / max(1, len(src_han))
        if miss or extra:
            problems.append((rel, len(miss), len(extra), ratio, miss[:2], extra[:2]))
    return problems


def main():
    ap = argparse.ArgumentParser(description='Word → Markdown 反向导入')
    ap.add_argument('-f', '--file', action='append', help='指定 docx（可多次）')
    ap.add_argument('--only', nargs='*', choices=FOLDERS, help='只处理部分分类文档')
    ap.add_argument('--stage', action='store_true', help='把还原出的 md 写到 Word回写/')
    ap.add_argument('--apply', action='store_true', help='直接写回 novel-project/')
    args = ap.parse_args()

    if not BASE.is_dir():
        raise SystemExit('找不到源目录 %s' % BASE)
    index = build_index()

    if args.file:
        targets = [Path(f) for f in args.file]
    else:
        want = args.only or FOLDERS
        targets = [SRC_DIR / ('%s-%s.docx' % (BOOK, cn)) for cn in want]
    targets = [t for t in targets if t.exists()]
    if not targets:
        raise SystemExit('Word导出/ 下没有可处理的 docx，先跑 tools/export_word.py')

    merged = {}
    news = {}
    titles = {}
    for t in targets:
        merged_all = '全量合集' in t.stem
        cn = next((f for f in FOLDERS if f in t.stem), None)
        level, prose = (3, False) if merged_all else (2, cn == '正文')
        r = Restorer(index, level, prose, folder=None if merged_all or not cn else cn)
        r.feed(Document(str(t)))
        merged.update(r.files)
        news.update(r.new_files)
        titles.update(r.new_titles)
        scope = FOLDERS if merged_all else ([cn] if cn else [])
        expect = {v for v in index.values() if v.split('/')[0] in scope}
        absent = sorted(expect - set(r.files))
        print('%-28s 还原 %d 个文件%s，Word 中新增 %d，源有但 Word 缺 %d'
              % (t.name, len(r.files), '/%d' % len(expect) if expect else '',
                 r.stats['新文件'], len(absent)))
        for a in absent[:5]:
            print('      缺: %s' % a)

    if news:
        print('\nWord 里新增、源目录中还没有的文件：')
        for rel in news:
            print('   + %-30s ← 标题「%s」' % (rel, titles.get(rel, rel)))

    problems = compare(index, merged)
    if problems:
        print('\n与源文件不一致（%d 个文件）：' % len(problems))
        for rel, nm, ne, ratio, miss, extra in problems[:15]:
            print('  %s  缺%d行/多%d行 字数比%.3f' % (rel, nm, ne, ratio))
            for l in miss:
                print('      源有·还原缺: %s' % l[:60])
            for l in extra:
                print('      还原多出的: %s' % l[:60])
        if len(problems) > 15:
            print('  ...另有 %d 个' % (len(problems) - 15))
    else:
        print('\n保真度比对：还原结果与全部源文件逐行一致')

    if not (args.stage or args.apply):
        print('\n（干跑模式，未写任何文件。加 --stage 出稿到 Word回写/，或 --apply 写回源目录）')
        return 0

    dest_root = STAGE_DIR if args.stage else BASE
    if args.apply:
        dirty = [rel for rel in merged if (BASE / rel).exists()]
        print('\n!! 将就地覆盖 %d 个源文件，建议先 `git commit` 存档' % len(dirty))
    if dest_root.exists() and args.stage:
        import shutil
        shutil.rmtree(dest_root)
    written = 0
    for rel, lines in list(merged.items()) + list(news.items()):
        out = dest_root / rel
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text('\n'.join(tidy(lines)).rstrip() + '\n', encoding='utf-8')
        written += 1
    print('\n已写出 %d 个 md → %s' % (written, dest_root))
    return 0


if __name__ == '__main__':
    sys.exit(main())
