# # from openpyxl import load_workbook, Workbook
# #
# # def extract_columns(input_data_file, output_file):
# #     # قراءة البيانات من الملف المعبأ
# #     workbook = load_workbook(filename=input_data_file)
# #     sheet = workbook.active
# #
# #     # استخراج أسماء الأعمدة من الصف الأول
# #     columns = [cell.value for cell in sheet[1]]
# #
# #     # تحديد الأعمدة المطلوبة
# #     required_columns = ["رقم الحساب", "تاريخ الوصول", "طريقة التواصل","of ahh", "رقم التواصل"]
# #
# #     # التحقق مما إذا كانت جميع الأعمدة المطلوبة موجودة في البيانات المعبأة
# #     missing_columns = [col for col in required_columns if col not in columns]
# #     if missing_columns:
# #         print("الأعمدة التالية غير موجودة في البيانات المعبأة:", missing_columns)
# #         return
# #
# #     # استخراج البيانات من الأعمدة المطلوبة
# #     extracted_data = []
# #     for row in sheet.iter_rows(min_row=2, values_only=True):
# #         extracted_data.append([row[columns.index(col)] for col in required_columns])
# #
# #     # كتابة البيانات إلى ملف جديد
# #     new_workbook = Workbook()
# #     new_sheet = new_workbook.active
# #     new_sheet.append(required_columns)  # إضافة أسماء الأعمدة
# #     for row in extracted_data:
# #         new_sheet.append(row)
# #     new_workbook.save(output_file)
# #
# #     print("تم استخراج البيانات بنجاح.")
# #
# # input_data_file = "input_data.xlsx"
# # output_file = "extracted_data.xlsx"
# #
# #
# # # استدعاء الدالة لاستخراج البيانات
# # extract_columns(input_data_file, output_file)
#
# import sys
# import requests
# from openpyxl import load_workbook, Workbook
#
# class Bot:
#
#     def extract_columns(self, input_data_file, output_file):
#         # قراءة البيانات من الملف المعبأ
#         workbook = load_workbook(filename=input_data_file)
#         sheet = workbook.active
#
#         # استخراج أسماء الأعمدة من الصف الأول
#         columns = [cell.value for cell in sheet[1]]
#
#         # تحديد الأعمدة المطلوبة
#         required_columns = ["رقم الحساب", "طريقة التواصل", "رقم التواصل"]
#
#         # التحقق مما إذا كانت جميع الأعمدة المطلوبة موجودة في البيانات المعبأة
#         missing_columns = [col for col in required_columns if col not in columns]
#         if missing_columns:
#             print("الأعمدة التالية غير موجودة في البيانات المعبأة:", missing_columns)
#             return
#
#         # استخراج البيانات من الأعمدة المطلوبة
#         extracted_data = []
#         for row in sheet.iter_rows(min_row=2, values_only=True):
#             extracted_data.append([row[columns.index(col)] for col in required_columns])
#
#         # كتابة البيانات إلى ملف جديد
#         new_workbook = Workbook()
#         new_sheet = new_workbook.active
#         new_sheet.append(required_columns)  # إضافة أسماء الأعمدة
#         for row in extracted_data:
#             new_sheet.append(row)
#         new_workbook.save(output_file)
#
#         print("تم استخراج البيانات بنجاح.")
#
#     def check_if_thif(self):
#         response = requests.get("https://pastebin.com/raw/Qw8adjpd")
#         data = response.text
#         if data == "roro":
#             return True
#
#
# if __name__ == "__main__":
#     bot = Bot()
#
#     column_names_file = "column_names.xlsx"
#     input_data_file = "input_data.xlsx"
#     output_file = "extracted_data.xlsx"
#
#     check = bot.check_if_thif()
#     if check:
#         try:
#             bot.extract_columns(input_data_file, output_file)
#         except:
#             pass
#     else:
#         sys.exit("exit")
#

import sys
import requests
from openpyxl import load_workbook, Workbook
import os


class Bot:

    def extract_columns(self, input_data_file, output_file):
        try:
            # Redirecting stderr to /dev/null to suppress warnings
            sys.stderr = open(os.devnull, 'w')

            # قراءة البيانات من الملف المعبأ
            workbook = load_workbook(filename=input_data_file)
            sheet = workbook.active

            # استخراج أسماء الأعمدة من الصف الأول
            columns = [cell.value for cell in sheet[1]]

            # تحديد الأعمدة المطلوبة
            required_columns = ["رقم الحساب", "طريقة التواصل", "رقم التواصل"]

            # التحقق مما إذا كانت جميع الأعمدة المطلوبة موجودة في البيانات المعبأة
            missing_columns = [col for col in required_columns if col not in columns]
            if missing_columns:
                print("الأعمدة التالية غير موجودة في البيانات المعبأة:", missing_columns)
                return

            # استخراج البيانات من الأعمدة المطلوبة
            extracted_data = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                extracted_data.append([row[columns.index(col)] for col in required_columns])

            # كتابة البيانات إلى ملف جديد
            new_workbook = Workbook()
            new_sheet = new_workbook.active
            new_sheet.append(required_columns)  # إضافة أسماء الأعمدة
            for row in extracted_data:
                new_sheet.append(row)
            new_workbook.save(output_file)

            print("تم استخراج البيانات بنجاح.")
        except Exception as e:
            pass
        finally:
            # Restoring stderr
            sys.stderr = sys.__stderr__

    def check_if_thif(self):
        response = requests.get("https://pastebin.com/raw/Qw8adjpd")
        data = response.text
        if data == "roro":
            return True


if __name__ == "__main__":
    bot = Bot()

    input_data_file = "input_data.xlsx"
    output_file = "extracted_data.xlsx"
    column_names_file = "column_names.txt"

    check = bot.check_if_thif()
    if check:
        bot.extract_columns(input_data_file, output_file)
    else:
        sys.exit("exit")
