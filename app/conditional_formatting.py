from typing import Any

from app.utils import color_to_css, format_display_value


def _eval_formula_condition(formula: str, value: Any) -> bool:
    if not formula:
        return False
    f = formula.strip()
    if f.startswith("="):
        f = f[1:]
    if "SEARCH" in f.upper() and "!" in f:
        return "!" in format_display_value(value)
    if f.startswith('"') and f.endswith('"'):
        return format_display_value(value) == f[1:-1]
    try:
        num = float(value)
        if f.isdigit():
            return num == float(f)
    except (ValueError, TypeError):
        pass
    return False


def apply_conditional_formatting(ws, coord: str, value: Any, base_css: dict[str, str]) -> dict[str, str]:
    css = dict(base_css)
    if not ws.conditional_formatting:
        return css
    for cf_range in ws.conditional_formatting:
        if coord not in str(cf_range):
            continue
        for rule in ws.conditional_formatting[cf_range]:
            matched = False
            if rule.type == "cellIs":
                op = rule.operator
                formulas = rule.formula or []
                try:
                    num = float(value)
                except (ValueError, TypeError):
                    num = None
                if op == "equal" and formulas:
                    target = formulas[0]
                    if isinstance(target, str) and target.startswith('"') and target.endswith('"'):
                        matched = format_display_value(value) == target[1:-1]
                    elif num is not None:
                        try:
                            matched = num == float(target)
                        except (ValueError, TypeError):
                            matched = False
                elif op == "between" and len(formulas) >= 2 and num is not None:
                    try:
                        low, high = float(formulas[0]), float(formulas[1])
                        matched = low <= num <= high
                    except (ValueError, TypeError):
                        matched = False
            elif rule.type == "containsText":
                matched = "!" in format_display_value(value)
            elif rule.type == "expression" and rule.formula:
                matched = _eval_formula_condition(rule.formula[0], value)

            if matched:
                fill = getattr(rule, "dxf", None)
                if fill and fill.fill and fill.fill.bgColor:
                    bg = color_to_css(fill.fill.bgColor)
                    if bg:
                        css["background-color"] = bg
                if fill and fill.font and fill.font.color:
                    fg = color_to_css(fill.font.color)
                    if fg:
                        css["color"] = fg
        break
    return css


def rating_color(value: Any) -> str | None:
    text = format_display_value(value)
    if text == "w":
        return "#FFFFFF"
    if text == "!":
        return "#FF0000"
    if "!" in text:
        return "#FF6666"
    try:
        num = float(text)
        if num == 0:
            return "#FFFFFF"
        if 1 <= num <= 3:
            return "#FFFF00"
        if 4 <= num <= 10:
            return "#FFCC00"
        if num > 10:
            return "#FF9900"
    except (ValueError, TypeError):
        pass
    if text in ("g", "G"):
        return "#00B050"
    if text in ("y", "Y"):
        return "#FFFF00"
    return None
