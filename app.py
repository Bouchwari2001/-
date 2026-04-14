import io
import math
import re
import zipfile
from datetime import date
from pathlib import Path

import arabic_reshaper
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from bidi.algorithm import get_display
from matplotlib import font_manager
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.font_manager import FontProperties
from PIL import Image


st.set_page_config(page_title="نظام بطاقات الجرد", page_icon="🏫", layout="centered")

st.markdown(
    """
    <style>
    html, body, [data-testid="stAppViewContainer"], [data-testid="stSidebar"] {
        direction: rtl;
    }
    .stApp, .stMarkdown, .stText, .stAlert, .stButton, .stDownloadButton, label {
        text-align: right;
    }
    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div,
    div[data-testid="stFileUploader"] section {
        direction: rtl;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


ARABIC_RE = re.compile(r"[\u0600-\u06FF]")
PAGE_SIZE = (8.27, 11.69)
ROWS_PER_PAGE = 15
ROOM_KEYWORDS = ["Salle", "قاعة", "مختبر", "Laboratoire"]
TABLE_COLUMNS = ["ملاحظات", "الحالة", "رقم الجرد", "العدد", "بيان التجهيز / الأثاث", "رت"]
TABLE_WIDTHS = [0.20, 0.13, 0.13, 0.08, 0.36, 0.10]

BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"
FIXED_HEADER_IMAGE_PATH = ASSETS_DIR / "fixed_header.png"

BODY_FONT_CANDIDATES = [
    ASSETS_DIR / "body-arabic.ttf",
    ASSETS_DIR / "Tajawal-Medium.ttf",
    ASSETS_DIR / "Amiri-Regular.ttf",
]
HEADER_FONT_CANDIDATES = [
    ASSETS_DIR / "moroccan-header.ttf",
    Path(r"C:\Windows\Fonts\trado.ttf"),
    Path(r"C:\Windows\Fonts\tradbdo.ttf"),
    Path(r"C:\Windows\Fonts\arabtype.ttf"),
]


def load_font(candidates):
    for candidate in candidates:
        try:
            if candidate.exists():
                font_manager.fontManager.addfont(str(candidate))
                return FontProperties(fname=str(candidate))
        except Exception:
            continue
    return None


BODY_FONT_PROP = load_font(BODY_FONT_CANDIDATES)
HEADER_FONT_PROP = load_font(HEADER_FONT_CANDIDATES) or BODY_FONT_PROP

if BODY_FONT_PROP is not None:
    plt.rcParams["font.family"] = BODY_FONT_PROP.get_name()


def rtl_text(value):
    text = "" if value is None else str(value)
    if not text.strip():
        return ""
    if not ARABIC_RE.search(text):
        return text
    return get_display(arabic_reshaper.reshape(text))


def text_kwargs(font_prop=None, size=None, weight=None):
    kwargs = {}
    if font_prop is not None:
        kwargs["fontproperties"] = font_prop
    if size is not None:
        kwargs["fontsize"] = size
    if weight is not None:
        kwargs["fontweight"] = weight
    return kwargs


def load_image_safe(source):
    try:
        path = Path(source)
        if path.exists():
            return Image.open(path).convert("RGBA")
    except TypeError:
        pass
    except Exception:
        return None

    try:
        if source is not None:
            return Image.open(source).convert("RGBA")
    except Exception:
        return None

    return None


def safe_filename(value):
    return re.sub(r'[<>:"/\\|?*]+', "_", str(value)).strip() or "room"


def detect_header_row(df_raw):
    for i in range(min(20, len(df_raw))):
        row_str = " ".join(str(x) for x in df_raw.iloc[i].values)
        if any(keyword in row_str for keyword in ROOM_KEYWORDS):
            return i
    return 0


def detect_rooms(columns):
    return [
        col
        for col in columns
        if any(keyword in str(col) for keyword in ROOM_KEYWORDS) or str(col).strip() in {"SVT", "PC"}
    ]


def detect_equipment_column(columns):
    preferred_keywords = ["بيان", "التجهيز", "الأثاث", "materiel", "matériel", "designation"]
    for col in columns:
        lowered = str(col).strip().lower()
        if any(keyword in lowered for keyword in preferred_keywords):
            return col
    return columns[0] if columns else None


def prepare_card_dataframe(df, room, equipment_col):
    room_counts = pd.to_numeric(df[room], errors="coerce").fillna(0)
    valid_items = df[room_counts > 0]

    # Detect optional enrichment columns in the uploaded file
    has_inv_num   = "رقم الجرد" in df.columns
    has_condition = "الحالة"    in df.columns
    has_remarks   = "ملاحظات"   in df.columns

    rows = []
    for idx, (_, row) in enumerate(valid_items.iterrows(), start=1):
        equipment_name = str(row.get(equipment_col, "")).strip()
        if equipment_name.lower() in {"nan", "", "none"}:
            continue

        count_value = pd.to_numeric(pd.Series([row[room]]), errors="coerce").fillna(0).iloc[0]
        rows.append(
            {
                "ملاحظات":              str(row["ملاحظات"]).strip()   if has_remarks   and str(row["ملاحظات"])   not in {"nan", "None"} else "",
                "الحالة":               str(row["الحالة"]).strip()    if has_condition and str(row["الحالة"])    not in {"nan", "None"} else "",
                "رقم الجرد":            str(row["رقم الجرد"]).strip() if has_inv_num   and str(row["رقم الجرد"]) not in {"nan", "None"} else "",
                "العدد":                int(count_value),
                "بيان التجهيز / الأثاث": equipment_name,
                "رت":                   idx,
            }
        )

    return pd.DataFrame(rows, columns=TABLE_COLUMNS)


def build_excel_template():
    buffer = io.BytesIO()

    sample_rows = [
        {"رقم الجرد": "INV-001", "بيان التجهيز / الأثاث": "طاولة",  "الحالة": "جديد",       "ملاحظات": "",              "Salle 01": 10, "Salle 02": 0,  "مختبر": 0, "SVT": 0, "PC": 0},
        {"رقم الجرد": "INV-002", "بيان التجهيز / الأثاث": "كرسي",   "الحالة": "مستعمل",     "ملاحظات": "بعضها تالف",    "Salle 01": 20, "Salle 02": 18, "مختبر": 0, "SVT": 0, "PC": 0},
        {"رقم الجرد": "INV-003", "بيان التجهيز / الأثاث": "سبورة",  "الحالة": "جيد",         "ملاحظات": "",              "Salle 01": 1,  "Salle 02": 1,  "مختبر": 0, "SVT": 0, "PC": 0},
        {"رقم الجرد": "INV-004", "بيان التجهيز / الأثاث": "حاسوب",  "الحالة": "معطل",        "ملاحظات": "بحاجة إصلاح",  "Salle 01": 0,  "Salle 02": 0,  "مختبر": 0, "SVT": 0, "PC": 12},
    ]
    template_df = pd.DataFrame(
        sample_rows,
        columns=["رقم الجرد", "بيان التجهيز / الأثاث", "الحالة", "ملاحظات", "Salle 01", "Salle 02", "مختبر", "SVT", "PC"],
    )

    instructions_df = pd.DataFrame(
        {
            "تعليمات": [
                "املأ اسم التجهيز في عمود: بيان التجهيز / الأثاث",
                "أدخل رقم الجرد الخاص بكل تجهيز في عمود: رقم الجرد  (اختياري)",
                "حدد حالة التجهيز في عمود: الحالة  — مثلاً: جديد / مستعمل / معطل  (اختياري)",
                "أضف أي ملاحظات إضافية في عمود: ملاحظات  (اختياري)",
                "ضع عدد كل تجهيز داخل العمود الخاص بالقاعة أو المختبر",
                "يمكنك إضافة أعمدة جديدة مثل Salle 03 أو قاعة 1 أو مختبر",
                "بعد التعبئة احفظ الملف ثم ارفعه داخل التطبيق لتوليد ملفات PDF",
            ]
        }
    )

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        template_df.to_excel(writer, sheet_name="الجرد", index=False)
        instructions_df.to_excel(writer, sheet_name="تعليمات", index=False)

    buffer.seek(0)
    return buffer.getvalue()


def draw_box(ax, x, y, w, h, lw=1.2):
    rect = plt.Rectangle((x, y), w, h, fill=False, linewidth=lw, edgecolor="black")
    ax.add_patch(rect)


def draw_header_band(ax):
    color = "#244b7a"
    ax.plot([0.05, 0.26], [0.94, 0.94], color=color, lw=2.2)
    ax.plot([0.74, 0.95], [0.94, 0.94], color=color, lw=2.2)
    ax.plot([0.26, 0.28], [0.94, 0.946], color=color, lw=2.2)
    ax.plot([0.28, 0.72], [0.946, 0.946], color=color, lw=2.2)
    ax.plot([0.72, 0.74], [0.946, 0.94], color=color, lw=2.2)


def get_page_rows(card_df, page_index):
    start = page_index * ROWS_PER_PAGE
    end = start + ROWS_PER_PAGE
    page_df = card_df.iloc[start:end].copy()

    rows = []
    for offset in range(ROWS_PER_PAGE):
        rank = start + offset + 1
        if offset < len(page_df):
            row = page_df.iloc[offset]
            rows.append(
                [
                    rtl_text(row["ملاحظات"]),
                    rtl_text(row["الحالة"]),
                    rtl_text(row["رقم الجرد"]),
                    str(row["العدد"]),
                    rtl_text(row["بيان التجهيز / الأثاث"]),
                    str(rank),
                ]
            )
        else:
            rows.append(["", "", "", "", "", str(rank)])
    return rows


def draw_default_header(ax, academy_name, directorate_name, school_name):
    draw_header_band(ax)
    ax.text(
        0.78,
        0.93,
        rtl_text(academy_name or "الأكاديمية الجهوية"),
        ha="right",
        va="top",
        **text_kwargs(HEADER_FONT_PROP, size=10.5, weight="bold"),
    )
    ax.text(
        0.78,
        0.904,
        rtl_text(directorate_name or "المديرية الإقليمية"),
        ha="right",
        va="top",
        **text_kwargs(HEADER_FONT_PROP, size=8.8),
    )
    ax.text(
        0.78,
        0.878,
        rtl_text(school_name or "اسم المؤسسة"),
        ha="right",
        va="top",
        **text_kwargs(HEADER_FONT_PROP, size=8.8),
    )


def draw_header(ax, header_mode, header_image, academy_name, directorate_name, school_name):
    if header_image is not None:
        extent = [0.05, 0.95, 0.84, 0.965] if header_mode == "fixed" else [0.05, 0.95, 0.845, 0.965]
        ax.imshow(header_image, extent=extent, aspect="auto", zorder=2)
    else:
        draw_default_header(ax, academy_name, directorate_name, school_name)

    info_y_start = 0.83 if header_image is not None else 0.835
    ax.text(
        0.5,
        info_y_start,
        rtl_text(academy_name),
        ha="center",
        va="top",
        **text_kwargs(HEADER_FONT_PROP, size=11.5, weight="bold"),
    )
    ax.text(
        0.5,
        info_y_start - 0.028,
        rtl_text(directorate_name),
        ha="center",
        va="top",
        **text_kwargs(HEADER_FONT_PROP, size=9.4),
    )
    ax.text(
        0.5,
        info_y_start - 0.054,
        rtl_text(school_name),
        ha="center",
        va="top",
        **text_kwargs(HEADER_FONT_PROP, size=9.4),
    )


def draw_table(ax, rows):
    col_labels = [rtl_text(col) for col in TABLE_COLUMNS]
    table = ax.table(
        cellText=rows,
        colLabels=col_labels,
        colWidths=TABLE_WIDTHS,
        cellLoc="center",
        bbox=[0.035, 0.16, 0.91, 0.36],
    )
    table.auto_set_font_size(False)
    table.set_fontsize(8.9)

    for (row, col), cell in table.get_celld().items():
        cell.set_linewidth(1.2)
        if row == 0:
            cell.set_facecolor("#d3d3d3")
            cell.set_text_props(weight="bold")
            cell.set_height(0.026)
        else:
            cell.set_height(0.025)

        if BODY_FONT_PROP is not None:
            cell.get_text().set_fontproperties(BODY_FONT_PROP)

        if col == 4:
            cell.get_text().set_ha("right")
        elif row > 0 and col in [0, 1, 2, 3, 5]:
            cell.get_text().set_ha("center")


def draw_page(ax, room, update_year, rows, header_mode, header_image, school_name, directorate_name, academy_name):
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)

    draw_header(ax, header_mode, header_image, academy_name, directorate_name, school_name)
    draw_table(ax, rows)

    ax.text(0.5, 0.735, "FICHE RECAPITULATIVE DE L'INVENTAIRE", ha="center", va="center", fontsize=12.5, fontweight="bold")
    ax.text(0.5, 0.695, rtl_text("بطاقة توطين المجرود"), ha="center", va="center", **text_kwargs(BODY_FONT_PROP, size=16, weight="bold"))

    draw_box(ax, 0.035, 0.61, 0.34, 0.035)
    ax.text(0.43, 0.627, rtl_text("المكان :"), ha="left", va="center", **text_kwargs(BODY_FONT_PROP, size=11.5, weight="bold"))
    ax.text(0.355, 0.627, rtl_text(room), ha="right", va="center", **text_kwargs(BODY_FONT_PROP, size=9.4))

    draw_box(ax, 0.035, 0.57, 0.34, 0.035)
    ax.text(0.43, 0.587, rtl_text("تاريخ التحيين :"), ha="left", va="center", **text_kwargs(BODY_FONT_PROP, size=11.5, weight="bold"))
    ax.text(0.355, 0.587, rtl_text(update_year), ha="right", va="center", **text_kwargs(BODY_FONT_PROP, size=9.4))

    ax.text(0.18, 0.135, rtl_text("توقيع رئيس المؤسسة"), ha="center", va="center", **text_kwargs(BODY_FONT_PROP, size=10))
    ax.text(0.72, 0.135, rtl_text("توقيع مسير المصالح المادية والمالية"), ha="center", va="center", **text_kwargs(BODY_FONT_PROP, size=10))


def build_room_pdf(room, card_df, update_year, header_mode, header_image, school_name, directorate_name, academy_name):
    pdf_buffer = io.BytesIO()
    total_pages = max(1, math.ceil(len(card_df) / ROWS_PER_PAGE))

    with PdfPages(pdf_buffer) as pdf:
        for page_index in range(total_pages):
            fig, ax = plt.subplots(figsize=PAGE_SIZE)
            fig.patch.set_facecolor("white")

            draw_page(
                ax=ax,
                room=room,
                update_year=update_year,
                rows=get_page_rows(card_df, page_index),
                header_mode=header_mode,
                header_image=header_image,
                school_name=school_name,
                directorate_name=directorate_name,
                academy_name=academy_name,
            )

            pdf.savefig(fig, bbox_inches="tight", pad_inches=0.2)
            plt.close(fig)

    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()


def read_inventory_file(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, header=None)
    header_idx = detect_header_row(df_raw)

    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, header=header_idx)
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
    return df


def render_header_controls():
    academy_name = st.text_input("🌍 الأكاديمية / الجهة", value="الأكاديمية الجهوية")
    directorate_name = st.text_input("🏢 المديرية", value="المديرية الإقليمية")
    school_name = st.text_input("🏫 اسم المؤسسة", value="ثانوية ألمدون الإعدادية")
    return academy_name, directorate_name, school_name


def render_header_mode():
    header_mode_label = st.radio(
        "🖼️ وضع الترويسة",
        options=["استعمال الترويسة الثابتة مع إضافة المعلومات", "رفع ترويسة خاصة"],
        index=0,
    )

    header_mode = "fixed" if header_mode_label == "استعمال الترويسة الثابتة مع إضافة المعلومات" else "custom"

    if header_mode == "fixed":
        header_image = load_image_safe(FIXED_HEADER_IMAGE_PATH)
        if header_image is not None:
            st.image(header_image, caption="الترويسة الثابتة المستعملة", use_container_width=True)
        else:
            st.warning("ملف الترويسة الثابتة غير موجود في assets/fixed_header.png. يمكنك إضافته أو اختيار رفع ترويسة خاصة.")
        return header_mode, header_image

    uploaded_header = st.file_uploader("🖼️ ارفع ترويسة خاصة", type=["png", "jpg", "jpeg"])
    header_image = load_image_safe(uploaded_header)
    if uploaded_header is not None and header_image is None:
        st.warning("تعذر قراءة ملف الترويسة. حاول رفع صورة بصيغة PNG أو JPG.")
    elif header_image is not None:
        st.image(header_image, caption="معاينة الترويسة الخاصة", use_container_width=True)

    return header_mode, header_image


def generate_all_pdfs(df, rooms, equipment_col, update_year, header_mode, header_image, school_name, directorate_name, academy_name):
    zip_buffer = io.BytesIO()
    generated = 0

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for room in rooms:
            card_df = prepare_card_dataframe(df, room, equipment_col)
            if card_df.empty:
                continue

            pdf_bytes = build_room_pdf(
                room=room,
                card_df=card_df,
                update_year=update_year,
                header_mode=header_mode,
                header_image=header_image,
                school_name=school_name,
                directorate_name=directorate_name,
                academy_name=academy_name,
            )
            zip_file.writestr(f"بطاقة_{safe_filename(room)}.pdf", pdf_bytes)
            generated += 1

    zip_buffer.seek(0)
    return generated, zip_buffer.getvalue()


st.title("🏫 تطبيق توليد بطاقات الجرد PDF")
st.markdown("ارفع ملف الجرد بصيغة Excel وسيتم توليد بطاقات PDF جاهزة للطباعة مع دعم العربية واتجاه RTL.")

workflow = st.radio(
    "🧭 اختر طريقة العمل",
    options=["رفع ملف جرد جاهز", "تحميل قالب Excel ثم تعبئته"],
    horizontal=True,
)

if workflow == "تحميل قالب Excel ثم تعبئته":
    st.info("يمكنك تحميل القالب، تعبئته ببيانات الجرد، ثم إعادة رفعه هنا لتوليد ملفات PDF.")
    st.download_button(
        label="📄 تحميل قالب Excel",
        data=build_excel_template(),
        file_name="inventory_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

uploaded_file = st.file_uploader("📥 ارفع ملف الجرد (Excel: xlsx أو xls)", type=["xlsx", "xls"])
current_year = date.today().year
default_update_year = f"{current_year}/{current_year + 1}"

academy_name, directorate_name, school_name = render_header_controls()
update_year = st.text_input("📅 تاريخ التحيين", value=default_update_year)
header_mode, header_image = render_header_mode()

if uploaded_file is not None:
    try:
        df = read_inventory_file(uploaded_file)
        rooms = detect_rooms(df.columns)
        equipment_default = detect_equipment_column(df.columns.tolist())

        if not rooms:
            st.warning("لم يتم العثور على أعمدة القاعات داخل الملف. تأكد من أسماء الأعمدة مثل Salle أو قاعة أو مختبر.")
        else:
            st.success(f"✅ تم التعرف على سطر العناوين الصحيح، وعدد القاعات المكتشفة هو {len(rooms)}.")

            equipment_col = st.selectbox(
                "🎯 اختر العمود الذي يحتوي على اسم التجهيز أو الأثاث",
                options=df.columns.tolist(),
                index=df.columns.tolist().index(equipment_default) if equipment_default in df.columns else 0,
            )

            if st.button("⚙️ توليد بطاقات PDF"):
                generated, zip_bytes = generate_all_pdfs(
                    df=df,
                    rooms=rooms,
                    equipment_col=equipment_col,
                    update_year=update_year,
                    header_mode=header_mode,
                    header_image=header_image,
                    school_name=school_name,
                    directorate_name=directorate_name,
                    academy_name=academy_name,
                )

                if generated:
                    st.success(f"🎉 تم إنشاء {generated} ملف PDF بنجاح.")
                    st.download_button(
                        label="📦 تحميل جميع البطاقات PDF",
                        data=zip_bytes,
                        file_name="inventory_cards_pdf.zip",
                        mime="application/zip",
                    )
                else:
                    st.warning("لم يتم إنشاء أي بطاقة. تحقق من القاعات والعمود المختار.")

    except Exception as exc:
        st.error(f"حدث خطأ أثناء معالجة الملف: {exc}")
        st.info("تأكد من أن ملف Excel منظم وأن أسماء الأعمدة واضحة.")
