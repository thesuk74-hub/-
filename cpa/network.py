"""주공정 경로 모식도 — 공정 네트워크(선후관계) 다이어그램. 주공정은 적색."""
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch  # noqa: E402

from . import gantt  # noqa: F401,E402  (한글 폰트 설정 공유)
from .schedule import topo_order  # noqa: E402


def render(res, path):
    acts = [a for a in res.acts]
    order, succ = topo_order(acts)
    A = {a.id: a for a in acts}
    level = {}
    for i in order:
        a = A[i]
        level[i] = 1 + max((level[r.pred] for r in a.preds), default=0)
    cols = {}
    for i in order:
        cols.setdefault(level[i], []).append(i)
    # 같은 열에서 주공정을 가운데(0행)에 배치
    pos = {}
    for lv, ids in cols.items():
        ids.sort(key=lambda x: (not A[x].critical, A[x].es))
        offs = [0, -1, 1, -2, 2, -3, 3, -4, 4]
        for k, i in enumerate(ids):
            pos[i] = (lv, offs[k] if k < len(offs) else k)
    nlv = max(level.values()) + 2
    ys = [y for _, y in pos.values()]
    fig_w = max(10, 1.75 * (nlv + 1))
    fig_h = max(3.2, 1.05 * (max(ys) - min(ys) + 1) + 1.2)
    fig, ax = plt.subplots(figsize=(min(fig_w, 22), fig_h))
    ax.set_xlim(-0.6, nlv + 0.6)
    ax.set_ylim(max(ys) + 0.8, min(ys) - 0.8)
    ax.axis("off")
    W, H = 0.8, 0.6

    def box(x, y, text, crit):
        ax.add_patch(FancyBboxPatch((x - W / 2, y - H / 2), W, H, boxstyle="round,pad=0.03",
                                    fc="#FFF5F5" if crit else "white", ec="#E53935" if crit else "#555555", lw=1.8 if crit else 1))
        ax.text(x, y, text, ha="center", va="center", fontsize=7.5, wrap=True)

    box(0, 0, f"착공\n준비기간 {res.prep}일", True)
    box(nlv, 0, f"정리기간 {res.cleanup}일\n준공", True)
    for i, (x, y) in pos.items():
        a = A[i]
        nm = a.name if len(a.name) <= 9 else a.name[:9] + "\n" + (a.name[9:18] + ("…" if len(a.name) > 18 else ""))
        box(x, y, f"{nm}\n{a.work:.0f}일", a.critical)

    def arrow(p0, p1, crit, label=""):
        (x0, y0), (x1, y1) = p0, p1
        ar = FancyArrowPatch((x0 + W / 2, y0), (x1 - W / 2, y1), arrowstyle="-|>", mutation_scale=10,
                             color="#E53935" if crit else "#777777", lw=1.6 if crit else 0.9,
                             connectionstyle=("arc3,rad=-0.35" if (y0 == y1 and x1 - x0 > 1) else
                                              "arc3,rad=0.0" if y0 == y1 else "angle3,angleA=0,angleB=90"))
        ax.add_patch(ar)
        if label:
            ax.text((x0 + x1) / 2, (y0 + y1) / 2 - 0.08, label, fontsize=6.5, color="#1565C0", ha="center")

    for i, (x, y) in pos.items():
        a = A[i]
        if not a.preds:
            arrow((0, 0), (x, y), a.critical)
        for r in a.preds:
            lab = "" if (r.type == "FS" and not r.lag) else r.label().split(":")[1]
            arrow(pos[r.pred], (x, y), a.critical and A[r.pred].critical, lab)
        if not succ[i]:
            arrow((x, y), (nlv, 0), a.critical)
    fig.tight_layout()
    fig.savefig(path, dpi=170)
    plt.close(fig)
