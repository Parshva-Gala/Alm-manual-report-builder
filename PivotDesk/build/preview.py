"""Render the Desk to PNG in both shipped and populated states."""
import os, sys
import assets, desk
from design_tokens import W, H, CANVAS
from shapes import to_html

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "preview")


def render_state(name, state, scale=2):
    shapes = desk.desk(state)
    desk.render_icons()
    page = to_html(shapes, W, H, CANVAS, assets.OUT)
    os.makedirs(OUT, exist_ok=True)
    html_path = os.path.join(OUT, "desk-%s.html" % name)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(page)
    png = os.path.join(OUT, "desk-%s.png" % name)
    assets.render(page, png, round(W * 4 / 3), round(H * 4 / 3), scale=scale, transparent=False)
    return png


if __name__ == "__main__":
    which = sys.argv[1:] or ["showcase", "empty", "tour"]
    for w in which:
        if w == "showcase":
            st = desk.showcase_state()
        elif w == "tour":
            # the tour on the step that explains the maturity gap
            st = desk.showcase_state()
            st.update({"toast": None, "tour": 5})
        else:
            st = desk.empty_state()
        print(render_state(w, st))
