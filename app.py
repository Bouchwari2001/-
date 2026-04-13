import streamlit as st
import pandas as pd
import io
import zipfile

# إعدادات التطبيق
st.set_page_config(page_title="نظام بطاقات الجرد", page_icon="🏫", layout="centered")

st.title("🏫 تطبيق توليد بطاقات توطين الجرد (Excel)")
st.markdown("قم برفع ملف الإكسيل الأصلي وسيقوم التطبيق بتوليد بطاقة إكسيل (.xlsx) منسقة وجاهزة للطباعة لكل قسم.")

# رفع ملف الإكسيل مباشرة
uploaded_file = st.file_uploader("📥 ارفع ملف الجرد (صيغة Excel - xlsx أو xls)", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # قراءة الملف مبدئياً للبحث عن السطر الذي يحتوي على القاعات (مثل Salle 01)
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        header_idx = 0
        # البحث في أول 20 سطر عن كلمة Salle أو قاعة
        for i in range(min(20, len(df_raw))):
            row_str = " ".join([str(x) for x in df_raw.iloc[i].values])
            if 'Salle' in row_str or 'قاعة' in row_str or 'مختبر' in row_str:
                header_idx = i
                break
        
        # قراءة الملف مرة أخرى باستخدام السطر الصحيح كـ Header
        df = pd.read_excel(uploaded_file, header=header_idx)
        
        # تنظيف أسماء الأعمدة من الأسطر الجديدة والمسافات
        df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]
        
        # استخراج أسماء القاعات الموجودة في الملف
        rooms = [col for col in df.columns if 'Salle' in col or 'قاعة' in col or 'مختبر' in col or col in ['SVT', 'PC']]
        
        st.success(f"✅ تم التعرف على السطر الصحيح و {len(rooms)} قاعة في الملف.")
        
        # اختيار عمود التجهيزات لتجنب أي أخطاء (Index out of bounds)
        all_columns = df.columns.tolist()
        equipment_col = st.selectbox("🎯 اختر العمود الذي يحتوي على 'بيان التجهيز / الأثاث':", all_columns)

        if st.button("⚙️ معالجة وتوليد بطاقات الإكسيل"):
            # تجهيز ملف ZIP في الذاكرة
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                cards_generated = 0
                
                for room in rooms:
                    # استخراج بيانات القاعة وتجاهل الفراغات
                    room_counts = pd.to_numeric(df[room], errors='coerce').fillna(0)
                    valid_items = df[room_counts > 0]
                    
                    if not valid_items.empty:
                        # تجهيز بيانات البطاقة
                        card_data = []
                        for idx, (_, row) in enumerate(valid_items.iterrows(), start=1):
                            equipment_name = str(row[equipment_col]).strip()
                            count = int(row[room])
                            
                            if equipment_name.lower() not in ["nan", "", "none"]:
                                card_data.append({
                                    "ملاحظات": "",
                                    "رقم الجرد": "",
                                    "العدد": count,
                                    "بيان التجهيز / الأثاث": equipment_name,
                                    "رت": idx
                                })
                        
                        if card_data:
                            df_card = pd.DataFrame(card_data)
                            
                            # إنشاء ملف إكسيل للقاعة وتنسيقه باستخدام xlsxwriter
                            excel_buffer = io.BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                                df_card.to_excel(writer, index=False, sheet_name='بطاقة الجرد', startrow=7)
                                
                                workbook = writer.book
                                worksheet = writer.sheets['بطاقة الجرد']
                                
                                # إعداد اتجاه الصفحة (من اليمين لليسار)
                                worksheet.right_to_left()
                                
                                # إنشاء التنسيقات
                                title_format = workbook.add_format({'bold': True, 'align': 'center', 'font_size': 14})
                                header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D3D3D3', 'border': 1})
                                cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
                                text_right_format = workbook.add_format({'bold': True, 'align': 'right'})
                                
                                # كتابة الترويسة العليا
                                worksheet.write('D1', 'المملكة المغربية\nوزارة التربية الوطنية والتعليم الأولي والرياضة\nثانوية ألمدون الإعدادية', text_right_format)
                                worksheet.merge_range('A4:E4', "FICHE RECAPITULATIVE DE L'INVENTAIRE", title_format)
                                worksheet.merge_range('A5:E5', "بطاقة توطين المجرود", title_format)
                                
                                worksheet.write('E7', f'المكان : {room}', text_right_format)
                                worksheet.write('B7', 'تاريخ التحيين : 2025/2026', text_right_format)
                                
                                # إعداد عرض الأعمدة ليناسب الطباعة
                                worksheet.set_column('A:A', 15) # ملاحظات
                                worksheet.set_column('B:B', 15) # رقم الجرد
                                worksheet.set_column('C:C', 10) # العدد
                                worksheet.set_column('D:D', 40) # بيان التجهيز
                                worksheet.set_column('E:E', 5)  # رت
                                
                                # تطبيق التنسيق على الجدول
                                for col_num, value in enumerate(df_card.columns.values):
                                    worksheet.write(7, col_num, value, header_format)
                                    
                                for row_num in range(len(df_card)):
                                    for col_num in range(len(df_card.columns)):
                                        worksheet.write(row_num + 8, col_num, df_card.iloc[row_num, col_num], cell_format)

                                # أماكن التوقيع أسفل الصفحة
                                last_row = len(df_card) + 10
                                worksheet.write(last_row, 0, 'توقيع رئيس المؤسسة', workbook.add_format({'bold': True, 'align': 'center'}))
                                worksheet.write(last_row, 3, 'توقيع مسير المصالح المادية والمالية', workbook.add_format({'bold': True, 'align': 'center'}))

                            # إضافة الملف إلى مجلد ZIP
                            safe_room_name = room.replace("/", "_").replace("\\", "_")
                            zip_file.writestr(f"بطاقة_{safe_room_name}.xlsx", excel_buffer.getvalue())
                            cards_generated += 1
                            
            if cards_generated > 0:
                st.success(f"🎉 تم بنجاح توليد وتنسيق {cards_generated} بطاقة بصيغة Excel!")
                st.download_button(
                    label="📦 تحميل جميع البطاقات (ملف مضغوط ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Inventory_Cards_Excel.zip",
                    mime="application/zip"
                )
            else:
                st.warning("لم يتم العثور على أي تجهيزات. تأكد من صحة العمود المختار.")
                
    except Exception as e:
        st.error(f"حدث خطأ: {e}")
        st.info("يرجى التأكد من رفع ملف الإكسيل الصحيح وأن البيانات منظمة بشكل سليم.")
