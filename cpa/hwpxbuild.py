"""HWPX(한글) 보고서 작성 도우미 — docbuild.Doc 과 같은 인터페이스(보고서 본문 코드를 그대로 사용).

python-hwpx 로 OWPML(.hwpx)을 작성한다. 머리글(사업명 / 장 제목), 바닥글(- 쪽번호 -), [표 n-m]·[그림 n-m] 캡션.
"""
import warnings

from hwpx import HwpxDocument

warnings.filterwarnings("ignore", category=DeprecationWarning, module=r"hwpx.*")
warnings.filterwarnings("ignore", category=DeprecationWarning)

FONT = "맑은 고딕"
NAVY = "#1F4E78"
HWPUNIT_PER_CM = 7200 / 2.54


def _fmt(v):
    if v is None:
        return ""
    if isinstance(v, float) and abs(v - round(v)) > 1e-9:
        return f"{v:,.1f}"
    if isinstance(v, (int, float)):
        return f"{int(round(v)):,}"
    return str(v)


class HwpxDoc:
    def __init__(self, project_name):
        self.d = HwpxDocument.new()
        self.pname = project_name
        self.ch = self.tn = self.fn = 0
        self.avail = 16.6
        self.sec = self.d.sections[0]
        self._runs = {}
        self.d.set_page_setup(width_mm=210, height_mm=297, orientation="PORTRAIT", margin_left_mm=22, margin_right_mm=22,
                              margin_top_mm=20, margin_bottom_mm=18, header_margin_mm=10, footer_margin_mm=10, section=self.sec)
        self._first = True

    # ── 내부 ──
    def _cp(self, size=10, bold=False, color=None, italic=False):
        k = (size, bold, color, italic)
        if k not in self._runs:
            self._runs[k] = self.d.styles.ensure_run(font=FONT, size=size, bold=bold, color=color or "#000000", italic=italic)
        return self._runs[k]

    def _p(self, text="", size=10, bold=False, color=None, align=None, indent=0.0, before=0, after=2, keep=False,
           first=None, italic=False, border=False):
        p = self.d.add_paragraph("", section=self.sec, include_run=False)
        if text:
            p.add_run(text, char_pr_id_ref=self._cp(size, bold, color, italic))
        kw = {}
        if align:
            kw["alignment"] = align
        if indent:
            kw["indent_left_mm"] = indent * 10
        if first:
            kw["first_line_indent_mm"] = first * 10
        if border:
            kw["bottom_border"], kw["border_color"], kw["border_width"] = True, NAVY, "0.4 mm"
        self.d.styles.apply_paragraph_format(paragraphs=[p], spacing_before_pt=before, spacing_after_pt=after,
                                             line_spacing_percent=150, keep_with_next=keep or None, **kw)
        return p

    # ── 머리글/구역 ──
    def first_header(self, right_text):
        self._header(self.sec, right_text)

    def set_header(self, section, right_text):
        self._header(self.sec, right_text)

    def _header(self, sec, text):
        self.d.set_header_text(f"{self.pname}    |    {text}", section=sec)
        self.d.set_page_number(position="BOTTOM_CENTER", prefix="- ", suffix=" -", section=sec)

    def new_section(self, header_text, landscape=False):
        self.sec = self.d.add_section()
        self.avail = 26.7 if landscape else 16.6
        if landscape:
            self.d.set_page_setup(width_mm=297, height_mm=210, orientation="LANDSCAPE", margin_left_mm=15, margin_right_mm=15,
                                  margin_top_mm=16, margin_bottom_mm=14, header_margin_mm=8, footer_margin_mm=8, section=self.sec)
        else:
            self.d.set_page_setup(width_mm=210, height_mm=297, orientation="PORTRAIT", margin_left_mm=22, margin_right_mm=22,
                                  margin_top_mm=20, margin_bottom_mm=18, header_margin_mm=10, footer_margin_mm=10, section=self.sec)
        self._header(self.sec, header_text)
        return self.sec

    # ── 본문 요소 ──
    def cover(self, title, sub, date_text, org, extra=None):
        for _ in range(6):
            self._p()
        self._p(title, 20, True, align="CENTER")
        self._p()
        self._p(sub, 28, True, NAVY, align="CENTER")
        if extra:
            self._p(extra, 11, color="#555555", align="CENTER")
        for _ in range(12):
            self._p()
        self._p(date_text, 16, True, align="CENTER")
        self._p(org, 16, True, align="CENTER")

    def toc(self, entries):
        self._p("< 목  차 >", 16, True, align="CENTER", after=10)
        for lvl, text in entries:
            self._p(text, 12 if lvl == 0 else 10.5, lvl == 0, indent=0.8 * lvl, after=6 if lvl == 0 else 2)

    def chapter(self, no, title, subs=(), title_page=True):
        self.ch, self.tn, self.fn = no, 0, 0
        self.new_section(f"제 {no} 장 {title}")
        for _ in range(8):
            self._p()
        self._p(f"제 {no} 장  {title}", 24, True, NAVY, border=True)
        self._p()
        for s in subs:
            self._p(s, 13, indent=1)
        if title_page:
            p = self._p(f"제{no}장 {title}", 16, True, align="CENTER")
            self.d.styles.apply_paragraph_format(paragraphs=[p], page_break_before=True)

    def h2(self, text):
        self._p(text, 13, True, NAVY, before=12, keep=True)

    def h3(self, text):
        self._p(text, 11.5, True, indent=0.3, before=8, keep=True)

    def h4(self, text):
        self._p(text, 10.5, True, indent=0.7, before=5, keep=True)

    def para(self, text, indent=1.2, size=10, bold=False, italic=False, color=None):
        c = None
        if color is not None:
            c = "#" + str(color).replace("#", "")[-6:]
        return self._p(text, size, bold, c, align="JUSTIFY", indent=indent, italic=italic)

    def bullet(self, text, indent=1.2, mark="-"):
        return self._p(f"{mark} {text}", 10, indent=indent + 0.4, first=-0.4)

    def note(self, text, indent=0.8):
        return self._p(text, 8.5, color="#444444", indent=indent)

    def boxed(self, lines, title=None, fill="F2F2F2"):
        t = self.d.add_table(1, 1, width=int(self.avail * HWPUNIT_PER_CM * 0.95), section=self.sec)
        cell = t.cell(0, 0)
        cell.set_text("")
        paras = []
        first = cell.paragraphs[0]
        if title:
            first.add_run(title, char_pr_id_ref=self._cp(10, True, NAVY))
            paras.append(first)
            first = None
        for ln in lines:
            p = first or cell.add_paragraph("")
            first = None
            p.add_run(ln, char_pr_id_ref=self._cp(10.5 if len(lines) == 1 else 9.5, len(lines) == 1))
            paras.append(p)
        t.set_cell_shading(0, 0, "#" + fill)
        self.d.styles.apply_paragraph_format(paragraphs=paras, alignment="CENTER" if len(lines) == 1 else "LEFT")
        self._p()

    def caption_table(self, title):
        self.tn += 1
        p = self._p("", align="CENTER", before=6, keep=True)
        p.add_run(f"[표 {self.ch}-{self.tn}]", char_pr_id_ref=self._cp(9.5, True))
        p.add_run(f"   {title}", char_pr_id_ref=self._cp(10, True))

    def table(self, header, rows, title=None, widths=None, size=8.5, align=None, hdr_fill="D9E2F3", bold_rows=(),
              fills=None, header_rows=None):
        if title:
            self.caption_table(title)
        hdrs = header_rows or [header]
        ncol = len(hdrs[0])
        nrow = len(hdrs) + len(rows)
        if not widths:
            lens = [1] * ncol
            for row in list(hdrs) + list(rows):
                for j, v in enumerate(row[:ncol]):
                    lens[j] = max(lens[j], min(len(_fmt(v)), 30))
            widths = [max(1.2, x) for x in lens]
        tot = sum(widths)
        if tot > self.avail:
            widths = [w * self.avail / tot for w in widths]
        tw = int(sum(widths) * HWPUNIT_PER_CM)
        t = self.d.add_table(nrow, ncol, width=tw, section=self.sec)
        t.set_column_widths([max(1, int(w * 100)) for w in widths])
        cps = {}
        hdr_p, cen, left, right = [], [], [], []
        for hi, hr in enumerate(hdrs):
            for j, h in enumerate(hr):
                c = t.cell(hi, j)
                c.set_text("")
                c.paragraphs[0].add_run(str(h), char_pr_id_ref=self._cp(size, True))
                hdr_p.append(c.paragraphs[0])
                t.set_cell_shading(hi, j, "#" + hdr_fill)
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                if j >= ncol:
                    break
                c = t.cell(len(hdrs) + i, j)
                c.set_text("")
                txt = _fmt(v)
                if txt:
                    k = (size, i in bold_rows)
                    if k not in cps:
                        cps[k] = self._cp(size, i in bold_rows)
                    c.paragraphs[0].add_run(txt, char_pr_id_ref=cps[k])
                a = align[j] if align else ("L" if isinstance(v, str) and len(v) > 14 else "C")
                {"L": left, "C": cen, "R": right}[a].append(c.paragraphs[0])
                if fills and fills[i]:
                    t.set_cell_shading(len(hdrs) + i, j, "#" + fills[i])
        for ps, al in ((hdr_p + cen, "CENTER"), (left, "LEFT"), (right, "RIGHT")):
            if ps:
                self.d.styles.apply_paragraph_format(paragraphs=ps, alignment=al, line_spacing_percent=110,
                                                     spacing_after_pt=0, spacing_before_pt=0)
        self._p(after=2)
        return t

    def figure(self, path, title, width_cm=16, max_h_cm=None):
        from PIL import Image
        w, h = Image.open(path).size
        if max_h_cm:
            width_cm = min(width_cm, max_h_cm * w / h)
        width_cm = min(width_cm, self.avail)
        with open(path, "rb") as f:
            data = f.read()
        self.d.add_picture(data, path.rsplit(".", 1)[-1].lower(), width_mm=width_cm * 10, height_mm=width_cm * 10 * h / w,
                           align="CENTER", section=self.sec)
        self.fn += 1
        self._p(f"[그림 {self.ch}-{self.fn}] {title}", 9.5, True, align="CENTER")

    def page_break(self):
        p = self._p()
        self.d.styles.apply_paragraph_format(paragraphs=[p], page_break_before=True)

    def save(self, path):
        self.d.save_to_path(path)
        _fix_package(path)


def _fix_package(path):
    """저장 후 보정: BinData 이미지의 manifest 항목 누락 보완, 작성자 메타데이터 정리 (mimetype 무압축 첫 항목 유지)."""
    import re
    import zipfile
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    names = [i.filename for i, _ in items]
    out = []
    for info, data in items:
        if info.filename == "Contents/content.hpf":
            x = data.decode("utf-8")
            for n in names:
                if n.startswith("BinData/") and f'href="{n}"' not in x:
                    bid = n.split("/")[-1].rsplit(".", 1)[0]
                    ext = n.rsplit(".", 1)[-1].lower()
                    mt = {"png": "image/png", "jpg": "image/jpg", "jpeg": "image/jpg", "bmp": "image/bmp"}.get(ext, "image/png")
                    x = x.replace("</opf:manifest>", f'<opf:item id="{bid}" href="{n}" media-type="{mt}" isEmbeded="1" /></opf:manifest>')
            x = re.sub(r"synthetic-fixture-author", "공사기간 적정성 검토", x)
            data = x.encode("utf-8")
        out.append((info, data))
    with zipfile.ZipFile(path, "w") as z:
        for info, data in out:
            ct = zipfile.ZIP_STORED if info.filename == "mimetype" else zipfile.ZIP_DEFLATED
            z.writestr(zipfile.ZipInfo(info.filename, date_time=info.date_time), data, compress_type=ct)
