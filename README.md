# Piping Comparison Script | اسکریپت مقایسه لوله‌کشی

This project includes a Python script designed to compare and match piping component descriptions between two Excel sheets.
The script extracts information such as size, type, material, degree, class, and other characteristics from each description
and compares them based on predefined criteria.

این پروژه شامل یک اسکریپت پایتون است که برای مقایسه و انطباق توصیفات اجزای لوله‌کشی بین دو شیت اکسل طراحی شده است.
اسکریپت اطلاعاتی مانند سایز، نوع، مواد، درجه، کلاس و سایر ویژگی‌ها را از توصیفات هر شیت استخراج کرده و آنها را مقایسه می‌کند.

**Support Email | ایمیل پشتیبانی**: hi.hosseinizaddi@gmail.com

## Requirements | پیش‌نیازها

To run this script, you need to install the following libraries:
برای اجرای این اسکریپت نیاز به نصب کتابخانه‌های زیر دارید:

- `pandas` - for Excel data processing | برای پردازش داده‌های اکسل
- `openpyxl` - for reading and saving Excel files | برای ذخیره و خواندن داده‌های اکسل
- `xlrd` - for reading older Excel file formats | برای خواندن فایل‌های اکسل با فرمت قدیمی

Use the following command to install:
برای نصب، دستور زیر را اجرا کنید:

```bash
pip install pandas openpyxl xlrd
#   i n v e n t o r y _ d e s c r i p t i o n _ m a t c h e r  
 


----------------------------------------------------------------------------------------------------------


# Usage Guide for Piping Comparison Script | راهنمای استفاده از اسکریپت مقایسه لوله‌کشی

1. Place the input Excel file in the `C:\Users\h.izadi\Desktop\piping\` directory.
   - Ensure that the file contains two sheets named `Sheet1` and `Sheet2`, each with a `Description` column.

   فایل اکسل ورودی را در مسیر `C:\Users\h.izadi\Desktop\piping\` قرار دهید.
   - اطمینان حاصل کنید که فایل شما شامل دو شیت به نام‌های `Sheet1` و `Sheet2` است که هر دو دارای ستونی به نام `Description` می‌باشند.

2. Run the script.
   - The script will call the following methods in sequence:
      - `display_comparison()`: shows matches in the console
      - `run_comparison()`: saves comparison results in the output file

   اسکریپت را اجرا کنید.
   - برنامه به ترتیب دو تابع زیر را فراخوانی می‌کند:
      - `display_comparison()`: نمایش تطابق‌ها در کنسول
      - `run_comparison()`: ذخیره نتایج مقایسه در فایل خروجی

3. The output Excel file will be saved in `C:\Users\h.izadi\Desktop\piping\out.xlsx`.

   فایل اکسل خروجی در مسیر `C:\Users\h.izadi\Desktop\piping\out.xlsx` ذخیره خواهد شد.

**Note**: To adjust comparison criteria, modify settings within the `compare_descriptions` method.

**توجه**: برای تغییر معیارهای مقایسه، تنظیمات تابع `compare_descriptions` را اصلاح کنید.

**Support Email | ایمیل پشتیبانی**: hi.hosseinizaddi@gmail.com
