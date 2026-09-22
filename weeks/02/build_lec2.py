# -*- coding: utf-8 -*-
"""Собрать Lec_2 в дизайне исходного Lec_1 (тема Simple Light, Lato, 16:9 10\")."""
from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parents[2]
FIG = Path(__file__).resolve().parent / "figures"
SRC = ROOT / "Презентации" / "Lec_1_intro2color.pptx"
OUT = ROOT / "Презентации" / "Lec_2_linear_model.pptx"

BLACK = RGBColor(0x00, 0x00, 0x00)
GRAY = RGBColor(0x59, 0x59, 0x59)
BLUE = RGBColor(0x42, 0x85, 0xF4)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PALE = RGBColor(0xEE, 0xEE, 0xEE)
FONT = "Lato"

# Как в исходном Google-деке: title (0.355", 0.463"), 9.19" x 0.70"
TITLE_L, TITLE_T = Inches(0.355), Inches(0.463)
TITLE_W, TITLE_H = Inches(9.19), Inches(0.70)


def delete_all_slides(prs: Presentation) -> None:
    sld_id_lst = prs.slides._sldIdLst
    ns = qn("r:id")
    for sld_id in list(sld_id_lst):
        r_id = sld_id.get(ns)
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)


def layout_named(prs: Presentation, name: str):
    return next(l for l in prs.slide_layouts if l.name == name)


def fill_body(placeholder, items):
    tf = placeholder.text_frame
    tf.clear()
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.level = 0
        p.text = item


def title_slide(prs, title, subtitle, foot):
    slide = prs.slides.add_slide(layout_named(prs, "TITLE"))
    slide.shapes.title.text = title
    # idx 1 = subtitle
    sub = slide.placeholders[1]
    sub.text = subtitle + "\n" + foot
    return slide


def section(prs, line1, line2=None):
    slide = prs.slides.add_slide(layout_named(prs, "SECTION_HEADER"))
    slide.shapes.title.text = line1 if not line2 else f"{line1}\n{line2}"
    return slide


def content(prs, title, items):
    slide = prs.slides.add_slide(layout_named(prs, "TITLE_AND_BODY"))
    slide.shapes.title.text = title
    fill_body(slide.placeholders[1], items)
    return slide


def _run(run, size, bold=False, color=BLACK, font=FONT, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font


def add_title(slide, text, size=24):
    box = slide.shapes.add_textbox(TITLE_L, TITLE_T, TITLE_W, TITLE_H)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    _run(r, size, bold=True)


def bullets(slide, items, top=1.25, left=0.45, width=9.1, size=16):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(3.9))
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(8)
        r = p.add_run()
        r.text = "•  " + item
        _run(r, size)


def pause(slide, text, top=4.85):
    box = slide.shapes.add_textbox(Inches(0.45), Inches(top), Inches(9.1), Inches(0.55))
    r = box.text_frame.paragraphs[0].add_run()
    r.text = text
    _run(r, 14, italic=True, color=BLUE)


def pic(slide, path, left, top, width=None, height=None):
    kw = {}
    if width:
        kw["width"] = Inches(width)
    if height:
        kw["height"] = Inches(height)
    slide.shapes.add_picture(str(path), Inches(left), Inches(top), **kw)


def formula_png(path: Path, tex: str, fontsize=26):
    fig, ax = plt.subplots(figsize=(8, 1.05))
    ax.axis("off")
    ax.text(0.5, 0.5, tex, ha="center", va="center", fontsize=fontsize)
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def make_figures():
    FIG.mkdir(exist_ok=True)
    lam = np.linspace(380, 730, 400)

    def gauss(mu, sig, a=1.0):
        return a * np.exp(-0.5 * ((lam - mu) / sig) ** 2)

    s, m, l = gauss(445, 22), gauss(540, 32), gauss(570, 38)
    cam_b, cam_g, cam_r = gauss(460, 18), gauss(530, 22), gauss(610, 28)
    fig, ax = plt.subplots(figsize=(6.4, 3.2))
    ax.plot(lam, s / s.max(), color="#4285F4", lw=2.0)
    ax.plot(lam, m / m.max(), color="#0F9D58", lw=2.0)
    ax.plot(lam, l / l.max(), color="#DB4437", lw=2.0, label="L, M, S")
    ax.plot(lam, cam_b / cam_b.max(), color="#4285F4", lw=1.4, ls="--")
    ax.plot(lam, cam_g / cam_g.max(), color="#0F9D58", lw=1.4, ls="--")
    ax.plot(lam, cam_r / cam_r.max(), color="#DB4437", lw=1.4, ls="--", label="камера R,G,B")
    ax.set_xlim(380, 730)
    ax.set_ylim(0, 1.08)
    ax.set_xlabel("λ, нм")
    ax.set_ylabel("относ. чувствительность")
    ax.legend(loc="upper right", fontsize=8, frameon=False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    fig.savefig(FIG / "camera_vs_lms.png", dpi=160, facecolor="white")
    plt.close(fig)

    from matplotlib.patches import FancyBboxPatch, Rectangle

    fig, ax = plt.subplots(figsize=(6.4, 2.5))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 4)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.add_patch(Rectangle((0.4, 0.5), 4.0, 2.6, facecolor="#EEEEEE", edgecolor="#000000", lw=1.4))
    ax.add_patch(Rectangle((5.6, 0.5), 4.0, 2.6, facecolor="#EEEEEE", edgecolor="#000000", lw=1.4))
    ax.text(2.4, 1.8, "тест  Φ", ha="center", va="center", fontsize=13)
    ax.text(7.6, 1.8, "смесь  αP₁+βP₂+γP₃", ha="center", va="center", fontsize=12)
    ax.annotate("", xy=(5.55, 1.8), xytext=(4.45, 1.8),
                arrowprops=dict(arrowstyle="<->", color="#4285F4", lw=1.6))
    ax.text(5.0, 3.45, "подогнать, пока поля неразличимы", ha="center", fontsize=11, color="#4285F4")
    fig.tight_layout()
    fig.savefig(FIG / "maxwell_match.png", dpi=160, facecolor="white")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(8.4, 1.85))
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 3)
    ax.axis("off")
    labels = ["I(λ)\nсвет", "R(λ)\nматериал", "Φ = I·R", "sᵢ(λ)\nприёмник", "cᵢ"]
    xs = [0.25, 2.55, 4.85, 7.15, 9.55]
    for x, lab in zip(xs, labels):
        ax.add_patch(FancyBboxPatch(
            (x, 0.55), 1.85, 1.7, boxstyle="round,pad=0.04,rounding_size=0.08",
            facecolor="#EEEEEE", edgecolor="#000000", lw=1.1,
        ))
        ax.text(x + 0.92, 1.4, lab, ha="center", va="center", fontsize=10)
    for x in xs[:-1]:
        ax.annotate("", xy=(x + 2.2, 1.4), xytext=(x + 2.05, 1.4),
                    arrowprops=dict(arrowstyle="->", color="#000000", lw=1.3))
    fig.tight_layout()
    fig.savefig(FIG / "pipeline.png", dpi=160, facecolor="white")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(7.2, 2.6))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 3.6)
    ax.axis("off")
    ax.add_patch(Rectangle((0.4, 0.35), 4.2, 2.9, facecolor="#FFFFFF", edgecolor="#000000", lw=1.1))
    ax.add_patch(Rectangle((5.4, 0.35), 4.2, 2.9, facecolor="#FFFFFF", edgecolor="#000000", lw=1.1))
    ax.text(2.5, 2.85, "диффузное", ha="center", fontsize=12, fontweight="bold")
    ax.text(7.5, 2.85, "блик (зеркальное)", ha="center", fontsize=12, fontweight="bold")
    ax.text(2.5, 1.5, "Φ ≈ I · R пигмента", ha="center", fontsize=11)
    ax.text(7.5, 1.5, "Φ ≈ спектр источника", ha="center", fontsize=11)
    fig.tight_layout()
    fig.savefig(FIG / "diffuse_specular.png", dpi=160, facecolor="white")
    plt.close(fig)

    formula_png(FIG / "eq_sensor.png", r"$c_i = \int s_i(\lambda)\,\Phi(\lambda)\,d\lambda$")
    formula_png(FIG / "eq_reflect.png", r"$\Phi(\lambda) = I(\lambda)\,R(\lambda)$")
    formula_png(FIG / "eq_full.png", r"$c_i = \int s_i(\lambda)\,I(\lambda)\,R(\lambda)\,d\lambda$")


def titled(prs, title):
    slide = prs.slides.add_slide(layout_named(prs, "TITLE_ONLY"))
    slide.shapes.title.text = title
    return slide


def build():
    make_figures()
    prs = Presentation(str(SRC))
    delete_all_slides(prs)

    title_slide(
        prs,
        "Вычислительная цветовая фотография",
        "Лекция 2. Линейная модель формирования",
        "IITP'26  ·  ИППИ РАН  ·  осень 2026",
    )

    content(prs, "Как устроен курс", [
        "До пары: материалы → вы готовите вопросы (это активность).",
        "На паре: повествование + разбор ваших вопросов.",
        "Тесты после лекций 3, 6, 9 и 12.",
        "Демо-семинар — сегодня, сразу после этой лекции.",
        "Оценка: 0.5 зачёт + 0.25 активность + 0.125 тесты + 0.125 явка + 0.4 практика.",
    ])
    content(prs, "Карта курса: 14 лекций", [
        "Блок А (25%) — колориметрия и зрительная система.",
        "Блок Б (75%) — репродукция и вычислительная фотография.",
        "Эта пара — линейная модель формирования изображений.",
        "Дальше в блоке А: свойства зрительной системы (пороги, адаптация, внешний вид).",
        "Не сегодня: XYZ-геометрия, CAM, ISP, sRGB.",
    ])
    s = content(prs, "Зачем модель", [
        "В файле нет цвета — есть числа после какого-то приёмника.",
        "Глаз и камера говорят на одном языке, только если ясно, что проинтегрировали.",
        "Дальше: трёхмерность, сложение равенств, материалы, приёмники, связь пространств.",
    ])
    pause(s, "Пауза. Цвет — свет, поверхность или наблюдатель? Сегодня — через интеграл.")

    section(prs, "Юнг–Максвелл")

    s = titled(prs, "Юнг: три приёмника")
    bullets(s, [
        "Большинство цветовых ощущений собирается из трёх независимых «основных».",
        "Это гипотеза о приёмниках, не о физике света: свет по-прежнему спектр.",
        "Три канала — про глаз, не про «красный / зелёный / синий свет».",
    ], width=5.15)
    pic(s, FIG / "lms_cones.png", 5.65, 1.35, width=4.15)

    s = titled(prs, "Максвелл: равенство, не имя цвета")
    bullets(s, [
        "Наблюдатель подгоняет смесь трёх основных под тестовое поле, пока не отличит.",
        "Измеряются коэффициенты равенства — не слово «зелёный».",
        "Двух основных мало; четырёх для равенства избыточно.",
    ], width=5.2, size=15)
    pic(s, FIG / "maxwell_match.png", 5.7, 1.35, width=4.1)
    pic(s, FIG / "maxwell_portrait.png", 6.85, 3.25, height=2.05)

    content(prs, "Что из опыта следует сразу", [
        "Трёхмерность равенств → три числа на спектр (пока без XYZ как стандарта).",
        "Разные спектры могут дать одно равенство — метамерия; лемма про ядро будет на Л3.",
        "Нужны законы: когда такие равенства можно складывать.",
    ])

    section(prs, "Грассман")

    content(prs, "Законы как правила равенств", [
        "Если A неотличим от B, перестановка полей равенства не ломает.",
        "Пропорциональность: смесь можно масштабировать.",
        "Аддитивность: равенства можно складывать.",
        "Непрерывность: близкие спектры → близкие равенства.",
        "Смысл: равенство — линейная операция на световых потоках.",
    ])

    s = titled(prs, "Формула, которая держит пару")
    pic(s, FIG / "eq_sensor.png", 1.2, 1.25, width=7.6)
    bullets(s, [
        "sᵢ — спектральная чувствительность i-го приёмника.",
        "Φ — спектр, который дошёл до приёмника.",
        "Глаз: i = L, M, S. Камера: i = R, G, B (или сколько фильтров есть).",
        "Если sᵢ фиксированы, Φ ↦ (cᵢ) линейно — это Грассман.",
    ], top=2.45)

    s = content(prs, "Где линейность кончается", [
        "Адаптация: приёмник «плывёт».",
        "Палочки и колбочки сразу (мезопическое).",
        "Очень яркое / очень тёмное.",
        "Материалы, где Φ ≠ I·R (флуоресценция — ниже).",
        "Порог различения ≠ равенство — следующие лекции блока А.",
    ])
    pause(s, "Пауза. Что в законах линейно и что уже нет?")

    section(prs, "Материалы")

    s = titled(prs, "Линейное отражение")
    pic(s, FIG / "eq_reflect.png", 1.3, 1.2, width=7.4)
    pic(s, FIG / "eq_full.png", 1.3, 2.15, width=7.4)
    pic(s, FIG / "pipeline.png", 0.55, 3.25, width=9.2)

    s = titled(prs, "Диффузное и зеркальное")
    pic(s, FIG / "diffuse_specular.png", 1.1, 1.25, width=7.9)
    bullets(s, [
        "Одно тело — два разных Φ в разных лучах. «Цвет объекта» неоднозначен.",
    ], top=4.15)

    content(prs, "Когда модель I·R ломается", [
        "Флуоресценция — поглотили одно λ, излучили другое; это не умножение спектров.",
        "Структурный / гониоцвет — R зависит от углов, не от «краски в точке».",
        "Просвечивающие среды — свет ходит внутри объёма.",
        "Линейный приёмник всё равно интегрирует то Φ, что дошло; врёт модель материала.",
    ])

    section(prs, "Приёмники: глаз и камера")

    content(prs, "Палочки", [
        "Один тип (родопсин), скотопическое зрение → нет цветовых равенств «как у колбочек».",
        "Не путать с третьим каналом камеры.",
        "В опытах дневной колориметрии их обычно выключают: яркость и ямка сетчатки.",
    ])

    s = titled(prs, "Колбочки L, M, S")
    bullets(s, [
        "Три спектральных приёмника: пики и перекрытие.",
        "Это не каналы R, G, B файла и не «красный / зелёный / синий свет».",
        "Три числа (c_L, c_M, c_S) — пространство колбочек.",
    ], width=4.9, size=15)
    pic(s, FIG / "lms_cones.png", 5.5, 1.3, width=4.3)

    content(prs, "Меланопсиновые клетки (ipRGC)", [
        "Ещё один фотопигмент, ганглиозные клетки.",
        "Зрачок, циркадные ритмы — не цветовое равенство Максвелла.",
        "В законы Грассмана про трихроматическое равенство не входят.",
        "«В глазу не ровно три приёмника» ≠ «цвет четырёхмерен в смысле Максвелла».",
    ])

    s = titled(prs, "Камера: те же интегралы, пока RAW линейный")
    bullets(s, [
        "cᵢ = ∫ sᵢ Φ dλ плюс шум и мозаика — вскользь.",
        "Фильтры Байера — конкретные s_R, s_G, s_B, не функции наблюдателя.",
        "Кривые другие → другие метамеры.",
    ], width=4.9, size=15)
    pic(s, FIG / "camera_vs_lms.png", 5.45, 1.3, width=4.35)

    s = titled(prs, "Глаз и камера рядом")
    rows = [
        ("", "Глаз (равенство)", "Камера (RAW)"),
        ("Что измеряем", "неразличимость полей", "числа на фотодиоде"),
        ("Сколько sᵢ", "три колбочки (в опыте)", "3 (Байер) или больше"),
        ("Линейность", "Грассман, с оговорками", "пока не тронули ISP"),
        ("Метамерия", "своя", "своя, другая"),
    ]
    table = s.shapes.add_table(len(rows), 3, Inches(0.45), Inches(1.3), Inches(9.1), Inches(3.4)).table
    table.columns[0].width = Inches(2.1)
    table.columns[1].width = Inches(3.5)
    table.columns[2].width = Inches(3.5)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = val
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    _run(run, 12, bold=(i == 0 or j == 0), color=BLACK)
            cell.fill.solid()
            cell.fill.fore_color.rgb = PALE if i == 0 else WHITE

    section(prs, "Связь пространств")

    s = titled(prs, "У каждого приёмника своё пространство")
    bullets(s, [
        "Спектр Φ один, векторы c разные: колбочки, камера A, камера B.",
        "«Связать пространства» — найти отображение c⁽ᴬ⁾ ↦ c⁽ᴮ⁾.",
    ], width=4.9, size=15)
    pic(s, FIG / "lettuce_metamers.png", 5.5, 1.3, width=4.3)

    content(prs, "Когда хватает матрицы 3×3", [
        "Оба набора sᵢ — линейные функционалы на одном и том же Φ.",
        "Если чувствительности камеры лежат в оболочке чувствительностей глаза — есть точная матрица.",
        "Иначе матрица — приближение; ошибка проявляется чужими метамерами.",
    ])

    s = content(prs, "Критерий Максвелла–Лютера–Айвса", [
        "Прибор измеряет «как глаз», если его sᵢ — линейная комбинация функций наблюдателя.",
        "Иначе у камеры и глаза разные классы метамеров.",
        "На Л10 это станет задачей CST. Сегодня — зачем сравнивать кривые.",
    ])
    pause(s, "Пауза. Можно ли пересчитать RAW в LMS одной матрицей? Когда да / когда нет.")

    section(prs, "Как измеряют s(λ)")

    content(prs, "Глаз: равенства по длинам волн", [
        "Монохроматический тест + три основных (Максвелл / Райт–Гилд → позже CMF).",
        "На каждой λ — три коэффициента в базисе выбранных основных.",
        "Это не кривые колбочек как есть; колбочки восстанавливают отдельно.",
        "Яркостная V(λ) — другой опыт. Геометрия — Л3.",
    ])
    content(prs, "Камера: известный спектр → линейный RAW", [
        "Монохроматор или узкие LED: Φ почти δ(λ), читаем RAW по каналам → оценка sᵢ(λ).",
        "Нужны линейность RAW, выключенные авто-WB и JPEG.",
        "Иначе измеряете ISP, не приёмник.",
    ])
    content(prs, "Два измерения — не одно и то же", [
        "Глаз: критерий — «поля совпали».",
        "Камера: критерий — «число в файле».",
        "CMF и кривые Canon живут в разных протоколах.",
        "Пороги Раутиана и аномалоскоп — не сегодня.",
    ])

    section(prs, "Закрытие")

    content(prs, "Что сегодня зафиксировали", [
        "Линейная модель c = ∫ s I R (когда R законна).",
        "Юнг–Максвелл: три приёмника и равенство.",
        "Грассман: линейность равенств.",
        "Материалы и приёмники задают, какой интеграл.",
        "Пространства разных s связаны не всегда матрицей.",
        "Чувствительности глаза и камеры измеряют по-разному.",
    ])
    content(prs, "Следующие пары блока А", [
        "Свойства зрительной системы дальше линейного равенства: пороги (Раутиан), адаптация, внешний вид.",
        "Геометрия пространства и лемма о метамерах — лекция 3.",
    ])
    content(prs, "Демо сегодня", [
        "Два спектра Φ, одни ответы c у глаза, другие — у «камеры» с чужими s.",
        "Метамерия чья: наблюдателя или прибора?",
        "Код не сдаём. После демо фиксируем, что входит в практику (0.4).",
    ])
    content(prs, "К зачёту", [
        "Карточки А1.3–4, А2.5–7, А2.9; измерение чувствительностей — А2.6.",
        "Раутиан (А2.8) — после лекции про пороги.",
        "Два вопроса из общей колоды, доля банка 25/75.",
    ])

    prs.save(str(OUT))


if __name__ == "__main__":
    build()
