from selenium_setting import create_driver, wait_for_elements, scroll_to_bottom
from selenium.webdriver.common.by import By
from openpyxl import Workbook
import streamlit as st
import pandas as pd
import time

# تنظیمات اولیه صفحه
st.set_page_config(
    page_title="استخراج لپ‌تاپ‌های دیجی‌کالا", page_icon="💻", layout="wide"
)

# عنوان برنامه
st.title("🛒 استخراج اطلاعات لپ‌تاپ‌های دیجی‌کالا")
st.markdown("---")

# توضیحات برنامه
with st.expander("راهنما و توضیحات برنامه"):
    st.write(
        """
    این برنامه اطلاعات لپ‌تاپ‌های موجود در دیجی‌کالا را استخراج می‌کند.
    
    - محدوده صفحات مورد نظر را انتخاب کنید
    - نوع خروجی را مشخص کنید (نمایش در برنامه یا ذخیره در اکسل)
    - دکمه شروع را بزنید
    
    **نکات مهم:**
    - برای عملکرد بهتر، مرورگر را نبندید
    - سرعت استخراج به سرعت اینترنت شما بستگی دارد
    - حداکثر 40 صفحه قابل استخراج است
    """
    )

# بخش تنظیمات
st.sidebar.header("تنظیمات استخراج")
min_page = st.sidebar.number_input(
    "شماره صفحه شروع", min_value=1, max_value=40, value=1
)
max_page = st.sidebar.number_input(
    "شماره صفحه پایان", min_value=1, max_value=40, value=3
)

output_options = st.sidebar.multiselect(
    "نوع خروجی", ["نمایش در برنامه", "ذخیره در اکسل"], default=["نمایش در برنامه"]
)

# دکمه شروع استخراج
if st.sidebar.button("شروع استخراج", type="primary"):
    if min_page > max_page:
        st.error("شماره صفحه شروع باید کوچکتر یا مساوی صفحه پایان باشد!")
    else:
        with st.spinner(f"در حال استخراج اطلاعات از صفحه {min_page} تا {max_page}..."):

            # نمایش نوار پیشرفت
            progress_bar = st.progress(0)
            status_text = st.empty()

            driver = create_driver(headless=True)
            all_names = []
            all_prices = []

            try:
                for i, page in enumerate(range(min_page, max_page + 1)):
                    try:
                        # به روزرسانی وضعیت
                        if min_page != max_page:
                            progress = int((i / (max_page - min_page)) * 100)
                            progress_bar.progress(progress)
                            status_text.text(
                                f"در حال پردازش صفحه {page} از {max_page}..."
                            )

                        # دریافت اطلاعات صفحه
                        link = f"https://www.digikala.com/search/notebook-netbook-ultrabook/?page={page}&price%5Bmin%5D=10000000&price%5Bmax%5D=700000000&q=%D9%84%D9%BE%20%D8%AA%D8%A7%D9%BE&sort=20"
                        driver.get(link)
                        driver.refresh()
                        driver.refresh()
                        driver.refresh()
                        scroll_to_bottom(driver)

                        # استخراج نام و قیمت‌ها
                        names = wait_for_elements(
                            driver,
                            By.CSS_SELECTOR,
                            'h3[class="ellipsis-2 text-body2-strong text-neutral-700 styles_VerticalProductCard__productTitle__6zjjN"]',
                            timeout=30,
                        )
                        prices = wait_for_elements(
                            driver,
                            By.CSS_SELECTOR,
                            'span[data-testid="price-final"]',
                            timeout=30,
                        )

                        # ذخیره اطلاعات
                        min_length = min(len(names), len(prices))
                        for j in range(min_length):
                            all_names.append(names[j].text)
                            all_prices.append(prices[j].text)

                        time.sleep(2)  # تاخیر برای جلوگیری از بلاک شدن

                    except Exception as e:
                        st.warning(f"خطا در پردازش صفحه {page}: {str(e)}")
                        continue

                # تکمیل نوار پیشرفت
                progress_bar.progress(100)
                status_text.text("استخراج با موفقیت انجام شد!")
                time.sleep(1)
                status_text.empty()

                # نمایش نتایج
                st.success(f"تعداد {len(all_names)} لپ‌تاپ استخراج شد")

                # ایجاد دیتافریم
                df = pd.DataFrame({"نام لپ‌تاپ": all_names, "قیمت (تومان)": all_prices})

                # نمایش نتایج در برنامه
                if "نمایش در برنامه" in output_options:
                    st.subheader("نتایج استخراج")
                    st.dataframe(df, height=600, use_container_width=True)

                # ذخیره در اکسل
                if "ذخیره در اکسل" in output_options:
                    with st.spinner("در حال ذخیره در فایل اکسل..."):
                        wb = Workbook()
                        ws = wb.active
                        ws.append(["نام لپ‌تاپ", "قیمت (تومان)"])

                        for name, price in zip(all_names, all_prices):
                            ws.append([name, price])

                        filename = f"laptops_digikala_page{min_page}-{max_page}.xlsx"
                        wb.save(filename)
                        st.success(f"اطلاعات با موفقیت در فایل {filename} ذخیره شد")

                        # دانلود فایل
                        with open(filename, "rb") as file:
                            st.download_button(
                                label="دانلود فایل اکسل",
                                data=file,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )

            finally:
                driver.quit()
                progress_bar.empty()
