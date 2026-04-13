import io
import math
import re
import zipfile
from datetime import date

import arabic_reshaper
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from bidi.algorithm import get_display
from matplotlib.backends.backend_pdf import PdfPages
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
TABLE_COLUMNS = ["ملاحظات", "رقم الجرد", "العدد", "بيان التجهيز / الأثاث", "رت"]
TABLE_WIDTHS = [0.18, 0.16, 0.1, 0.46, 0.1]


def rtl_text(value):
    text = "" if value is None else str(value)
    if not text.strip():
        return ""
    if not ARABIC_RE.search(text):
        return text
    return get_display(arabic_reshaper.reshape(text))


def safe_filename(value):
    return re.sub(r'[<>:"/\\|?*]+', "_", str(value)).strip() or "room"


def detect_header_row(df_raw):
    for i in range(min(20, len(df_raw))):
        row_str = " ".join(str(x) for x in df_raw.iloc[i].values)
        if any(keyword in row_str for keyword in ["Salle", "قاعة", "مختبر", "Laboratoire"]):
            return i
    return 0


def detect_rooms(columns):
    room_keywords = ["Salle", "قاعة", "مختبر", "Laboratoire"]
    return [
        col
        for col in columns
        if any(keyword in str(col) for keyword in room_keywords) or str(col).strip() in ["SVT", "PC"]
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

    rows = []
    for idx, (_, row) in enumerate(valid_items.iterrows(), start=1):
        equipment_name = str(row.get(equipment_col, "")).strip()
        if equipment_name.lower() in {"nan", "", "none"}:
            continue

        count_value = pd.to_numeric(pd.Series([row[room]]), errors="coerce").fillna(0).iloc[0]
        rows.append(
            {
                "ملاحظات": "",
                "رقم الجرد": "",
                "العدد": int(count_value),
                "بيان التجهيز / الأثاث": equipment_name,
                "رت": idx,
            }
        )

    return pd.DataFrame(rows, columns=TABLE_COLUMNS)


def build_excel_template():
    buffer = io.BytesIO()

    sample_rows = [
        {"بيان التجهيز / الأثاث": "طاولة", "Salle 01": 10, "Salle 02": 0, "مختبر": 0, "SVT": 0, "PC": 0},
        {"بيان التجهيز / الأثاث": "كرسي", "Salle 01": 20, "Salle 02": 18, "مختبر": 0, "SVT": 0, "PC": 0},
        {"بيان التجهيز / الأثاث": "سبورة", "Salle 01": 1, "Salle 02": 1, "مختبر": 0, "SVT": 0, "PC": 0},
        {"بيان التجهيز / الأثاث": "حاسوب", "Salle 01": 0, "Salle 02": 0, "مختبر": 0, "SVT": 0, "PC": 12},
    ]
    template_df = pd.DataFrame(sample_rows)

    instructions_df = pd.DataFrame(
        {
            "تعليمات": [
                "املأ اسم التجهيز في عمود: بيان التجهيز / الأثاث",
                "ضع عدد كل تجهيز داخل العمود الخاص بالقاعة أو الجناح",
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
                    rtl_text(row["رقم الجرد"]),
                    str(row["العدد"]),
                    rtl_text(row["بيان التجهيز / الأثاث"]),
                    str(rank),
                ]
            )
        else:
            rows.append(["", "", "", "", str(rank)])
    return rows


def draw_page(ax, room, school_name, update_year, rows, logo_image):
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)

    if logo_image is not None:
        ax.imshow(logo_image, extent=[0.05, 0.95, 0.86, 0.96], aspect="auto", zorder=3)
    else:
        draw_header_band(ax)
        ax.text(
            0.62,
            0.905,
            rtl_text("المملكة المغربية"),
            ha="right",
            va="top",
            fontsize=9,
        )
        ax.text(
            0.62,
            0.885,
            rtl_text("وزارة التربية الوطنية والتعليم الأولي والرياضة"),
            ha="right",
            va="top",
            fontsize=7.8,
        )
        ax.text(
            0.62,
            0.865,
            rtl_text(school_name),
            ha="right",
            va="top",
            fontsize=7.8,
        )

    col_labels = [rtl_text(col) for col in TABLE_COLUMNS]
    table = ax.table(
        cellText=rows,
        colLabels=col_labels,
        colWidths=TABLE_WIDTHS,
        cellLoc="center",
        bbox=[0.035, 0.195, 0.91, 0.51],
    )
    table.auto_set_font_size(False)
    table.set_fontsize(10.5)

    for (row, col), cell in table.get_celld().items():
        cell.set_linewidth(1.2)
        if row == 0:
            cell.set_facecolor("#d3d3d3")
            cell.set_text_props(weight="bold")
            cell.set_height(0.034)
        else:
            cell.set_height(0.037)
        if col == 3:
            cell.get_text().set_ha("right")
        if row > 0 and col in [0, 1, 2, 4]:
            cell.get_text().set_ha("center")

    ax.text(0.5, 0.845, "FICHE RECAPITULATIVE DE L'INVENTAIRE", ha="center", va="center", fontsize=12.5, fontweight="bold")
    ax.text(0.5, 0.812, rtl_text("بطاقة توطين المجرود"), ha="center", va="center", fontsize=17, fontweight="bold")

    draw_box(ax, 0.035, 0.738, 0.34, 0.035)
    ax.text(0.43, 0.755, rtl_text("المكان :"), ha="left", va="center", fontsize=11.5, fontweight="bold")
    ax.text(0.36, 0.755, rtl_text(room), ha="right", va="center", fontsize=9.8)

    draw_box(ax, 0.035, 0.700, 0.34, 0.035)
    ax.text(0.43, 0.717, rtl_text("تاريخ التحيين :"), ha="left", va="center", fontsize=11.5, fontweight="bold")
    ax.text(0.36, 0.717, rtl_text(update_year), ha="right", va="center", fontsize=9.8)

    ax.text(0.2, 0.185, rtl_text("توقيع رئيس المؤسسة"), ha="center", va="center", fontsize=11.5, fontweight="bold")
    ax.text(0.73, 0.185, rtl_text("توقيع مسير المصالح المادية و المالية"), ha="center", va="center", fontsize=11.5, fontweight="bold")


def build_room_pdf(room, card_df, school_name, update_year, logo_image):
    pdf_buffer = io.BytesIO()
    total_pages = max(1, math.ceil(len(card_df) / ROWS_PER_PAGE))

    with PdfPages(pdf_buffer) as pdf:
        for page_index in range(total_pages):
            fig, ax = plt.subplots(figsize=PAGE_SIZE)
            fig.patch.set_facecolor("white")
            rows = get_page_rows(card_df, page_index)
            draw_page(
                ax=ax,
                room=room,
                school_name=school_name,
                update_year=update_year,
                rows=rows,
                logo_image=logo_image,
            )
            pdf.savefig(fig, bbox_inches="tight", pad_inches=0.2)
            plt.close(fig)

    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()


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
school_name = st.text_input("🏫 اسم المؤسسة", value="ثانوية ألمدون الإعدادية")
update_year = st.text_input("📅 تاريخ التحيين", value=default_update_year)
logo_file = st.file_uploader("🖼️ شعار المؤسسة أو الترويسة الرسمية (اختياري)", type=["png", "jpg", "jpeg"])

logo_image = None
if logo_file is not None:
    try:
        logo_image = Image.open(logo_file).convert("RGBA")
        st.image(logo_image, caption="معاينة الشعار أو الترويسة", use_container_width=True)
    except Exception:
        st.warning("تعذر قراءة ملف الشعار. حاول رفع صورة بصيغة PNG أو JPG.")

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        header_idx = detect_header_row(df_raw)

        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=header_idx)
        df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

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
                            school_name=school_name,
                            update_year=update_year,
                            logo_image=logo_image,
                        )
                        zip_file.writestr(f"بطاقة_{safe_filename(room)}.pdf", pdf_bytes)
                        generated += 1

                if generated:
                    zip_buffer.seek(0)
                    st.success(f"🎉 تم إنشاء {generated} ملف PDF بنجاح.")
                    st.download_button(
                        label="📦 تحميل جميع البطاقات PDF",
                        data=zip_buffer.getvalue(),
                        file_name="inventory_cards_pdf.zip",
                        mime="application/zip",
                    )
                else:
                    st.warning("لم يتم إنشاء أي بطاقة. تحقق من القاعات والعمود المختار.")

    except Exception as exc:
        st.error(f"حدث خطأ أثناء معالجة الملف: {exc}")
        st.info("تأكد من أن ملف Excel منظم وأن أسماء الأعمدة واضحة.")
