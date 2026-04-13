import streamlit as st
import pandas as pd
import zipfile
import io

# App Configuration
st.set_page_config(page_title="نظام بطاقات الجرد", page_icon="🏫", layout="centered")

# Header
st.title("🏫 تطبيق توليد بطاقات توطين الجرد")
st.markdown("""
هذا التطبيق مخصص لتسهيل عمل الإدارة التربوية. 
فقط قم برفع ملف الإكسيل (CSV) المستخرج من نظام الجرد، وسيقوم التطبيق بتوليد بطاقات التوطين لكل قاعة تلقائياً في ملف مضغوط جاهز للطباعة.
""")

# File Uploader
uploaded_file = st.file_uploader("📥 ارفع ملف الجرد (صيغة CSV)", type=['csv'])

# The HTML Template (Built into the app so users don't have to upload two files)
html_template = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
 <head>
  <meta charset="utf-8"/>
  <style>
      body { font-family: Arial, sans-serif; margin: 40px; }
      table { width: 100%; border-collapse: collapse; margin-top: 20px; text-align: center; }
      th, td { border: 1px solid black; padding: 8px; }
      th { background-color: #cccccc; }
      .header-text { text-align: center; font-weight: bold; }
      .signatures { display: flex; justify-content: space-between; margin-top: 50px; font-weight: bold; }
  </style>
 </head>
 <body>
  <div style="text-align: right;">
      <p>المملكة المغربية<br>وزارة التربية الوطنية والتعليم الأولي والرياضة<br>ثانوية ألمدون الإعدادية</p>
  </div>
  <h1><p style="text-align: center;"><b>FICHE RECAPITULATIVE DE L'INVENTAIRE</b></p></h1>
  <h2><h2 style="text-align: center;">بطاقة توطين المجرود</h2></h2>
  
  <table style="border: none; margin-bottom: 20px;">
   <tr style="border: none;">
    <td style="border: none; width: 20%; text-align: right; font-weight: bold;">المكان :</td>
    <td style="border: none; width: 30%; text-align: right;">{room_name}</td>
    <td style="border: none; width: 20%; text-align: right; font-weight: bold;">الجناح :</td>
    <td style="border: none; width: 30%; text-align: right;">التعليم العام / العلمي</td>
   </tr>
   <tr style="border: none;">
    <td style="border: none; text-align: right; font-weight: bold;">تاريخ التحيين :</td>
    <td style="border: none; text-align: right;">2025/2026</td>
    <td style="border: none;"></td>
    <td style="border: none;"></td>
   </tr>
  </table>

  <table>
   <thead>
    <tr>
     <th style="width: 25%;">ملاحظات</th>
     <th style="width: 15%;">رقم الجرد</th>
     <th style="width: 10%;">العدد</th>
     <th style="width: 40%;">بيان التجهيز / الأثاث</th>
     <th style="width: 10%;">رت</th>
    </tr>
   </thead>
   <tbody>
    {table_rows}
   </tbody>
  </table>

  <div class="signatures">
      <div><u>توقيع مسير المصالح المادية والمالية</u></div>
      <div><u>توقيع رئيس المؤسسة</u></div>
  </div>
 </body>
</html>
"""

if uploaded_file is not None:
    try:
        # Read the uploaded CSV
        df = pd.read_csv(uploaded_file, header=11)
        df.columns = [str(col).replace('\n', ' ').strip() for col in df.columns]

        # Target rooms
        rooms = [
            "Salle 01", "Salle 02", "Salle 03", "Salle 04", "Salle 05", "Salle 06", 
            "Salle 07", "Salle 08", "Salle 09", "Salle 10", "Salle 11", "Salle 12",
            "SVT", "PC", "قاعة الأساتذة", "قاعة جيني", 
            "مختبر الجناح العلمي", "مختبر الخاص بمادة بالاجتماعيات"
        ]
        
        equipment_col = df.columns[20]

        if st.button("⚙️ معالجة وتوليد البطاقات"):
            # Create a zip file in memory
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                cards_generated = 0
                
                for room in rooms:
                    if room in df.columns:
                        room_counts = pd.to_numeric(df[room], errors='coerce').fillna(0)
                        valid_items = df[room_counts > 0]
                        
                        if not valid_items.empty:
                            rows_html = ""
                            for index, (_, row) in enumerate(valid_items.iterrows(), start=1):
                                equipment_name = str(row[equipment_col]).strip()
                                count = int(row[room])
                                
                                if equipment_name.lower() not in ["nan", "", "none"]:
                                    rows_html += f"<tr><td></td><td></td><td>{count}</td><td>{equipment_name}</td><td>{index}</td></tr>"
                            
                            if rows_html != "":
                                final_html = html_template.format(room_name=room, table_rows=rows_html)
                                safe_room_name = room.replace("/", "_").replace("\\", "_")
                                file_name = f"Card_{safe_room_name}.html"
                                
                                # Write HTML directly to the zip file buffer
                                zip_file.writestr(file_name, final_html.encode('utf-8'))
                                cards_generated += 1
            
            st.success(f"✅ تمت العملية بنجاح! تم إنشاء {cards_generated} بطاقة جرد.")
            
            # Show download button for the generated ZIP
            st.download_button(
                label="📦 تحميل جميع البطاقات (ملف مضغوط ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Inventory_Cards_AlMadoun.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة الملف. يرجى التأكد من أن الملف مطابق للنموذج المطلوب. التفاصيل: {e}")
