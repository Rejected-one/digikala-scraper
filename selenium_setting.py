import requests
import subprocess
import os
import tempfile
import time
import json
from pathlib import Path
import zipfile
import winreg
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from typing import List, Optional
import random
import sys
from kill_browser import  kill_process

def get_current_edge_version():
    """
    نسخه فعلی مرورگر Edge را از رجیستری یا فایل اجرایی دریافت می‌کند.

    Returns:
        str or None: نسخه مرورگر یا None در صورت خطا یا عدم وجود.
    """
    try:
        import winreg

        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Edge\BLBeacon"
        )
        version = winreg.QueryValueEx(key, "version")[0]
        winreg.CloseKey(key)
        return version
    except:
        try:
            result = subprocess.run(
                [
                    "powershell",
                    "-Command",
                    '(Get-Item "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe").VersionInfo.ProductVersion',
                ],
                capture_output=True,
                text=True,
                timeout=10,
            )

            if result.returncode == 0 and result.stdout.strip():
                return result.stdout.strip()
        except:
            pass
    return None


def get_latest_edge_version():
    """
    آخرین نسخه مرورگر Edge را از API رسمی مایکروسافت دریافت می‌کند.

    Returns:
        dict or None: دیکشنری شامل نسخه و آدرس دانلود یا None در صورت خطا.
    """
    try:
        url = "https://edgeupdates.microsoft.com/api/products"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()

        data = response.json()

        # پیدا کردن نسخه Stable برای Windows x64
        for product in data:
            if product.get("Product") == "Stable":
                releases = product.get("Releases", [])
                for release in releases:
                    if (
                        release.get("Platform") == "Windows"
                        and release.get("Architecture") == "x64"
                    ):
                        return {"version": release.get("ProductVersion"), "url": None}

        return None
    except Exception as e:
        print(f"خطا در دریافت اطلاعات نسخه: {e}")
        return None


def download_edge_installer():
    """
    نصب‌کننده مرورگر Edge را برای ویندوز x64 دانلود می‌کند.

    Returns:
        str or None: مسیر فایل نصب‌کننده دانلود شده یا None در صورت خطا.
    """
    try:
        print("در حال دریافت اطلاعات دانلود...")
        url = "https://edgeupdates.microsoft.com/api/products"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()

        data = response.json()

        # پیدا کردن لینک دانلود برای Windows x64 Stable
        for product in data:
            if product.get("Product") == "Stable":
                releases = product.get("Releases", [])
                for release in releases:
                    if (
                        release.get("Platform") == "Windows"
                        and release.get("Architecture") == "x64"
                    ):
                        download_url = None
                        artifacts = release.get("Artifacts", [])
                        for artifact in artifacts:
                            if artifact.get("Location") == "Public":
                                download_url = artifact.get("Url")
                                break

                        if download_url:
                            print("در حال دانلود نصب‌کننده Edge...")
                            installer_response = requests.get(
                                download_url, headers=headers, timeout=120
                            )  # زمان بیشتر برای دانلود بزرگ
                            installer_response.raise_for_status()

                            # ذخیره فایل در پوشه موقت
                            temp_dir = tempfile.gettempdir()
                            installer_path = os.path.join(
                                temp_dir, "MicrosoftEdgeSetup.exe"
                            )

                            with open(installer_path, "wb") as f:
                                f.write(installer_response.content)

                            print(f"نصب‌کننده دانلود شد: {installer_path}")
                            return installer_path
        return None

    except Exception as e:
        print(f"خطا در دانلود نصب‌کننده: {e}")
        return None


def update_edge():
    """
    مرورگر Edge را به آخرین نسخه بروزرسانی می‌کند.

    Returns:
        bool: True اگر بروزرسانی موفق بود، False در غیر این صورت.
    """
    print("🔍 بررسی نسخه فعلی Edge...")
    current_version = get_current_edge_version()

    if current_version:
        print(f"نسخه فعلی: {current_version}")
    else:
        print("⚠️ نسخه فعلی پیدا نشد")

    print("🔍 در حال بررسی آخرین نسخه...")
    latest_info = get_latest_edge_version()

    if latest_info and latest_info["version"]:
        latest_version = latest_info["version"]
        print(f"آخرین نسخه: {latest_version}")

        # مقایسه نسخه‌ها
        if current_version and current_version == latest_version:
            print("✅ مرورگر Edge به‌روز است")
            return True
        elif current_version:
            print(f"🔄 نسخه جدیدی موجود است: {current_version} → {latest_version}")
        else:
            print("🔄 شروع به نصب آخرین نسخه...")
    else:
        print("⚠️ آخرین نسخه پیدا نشد")

    # دانلود و نصب
    print("📥 در حال دانلود نصب‌کننده...")
    installer_path = download_edge_installer()

    if not installer_path:
        print("❌ خطا در دانلود نصب‌کننده")
        return False

    try:
        # بستن Edge اگر باز باشد
        print("⏹️ در حال بستن Edge (در صورت باز بودن)...")
        subprocess.run(
            ["taskkill", "/F", "/IM", "msedge.exe"], capture_output=True, timeout=10
        )
        time.sleep(3)

        # نصب به‌روزرسانی
        print("🚀 در حال نصب به‌روزرسانی...")
        result = subprocess.run(
            [installer_path, "--silent", "--system-level", "--do-not-launch-chrome"],
            capture_output=True,
            timeout=300,
        )

        # حذف فایل نصب‌کننده
        try:
            os.remove(installer_path)
        except:
            pass

        if result.returncode == 0:
            print("✅ بروزرسانی با موفقیت انجام شد")
            # نمایش نسخه جدید
            new_version = get_current_edge_version()
            if new_version:
                print(f"نسخه جدید: {new_version}")
            return True
        else:
            error_msg = (
                result.stderr.decode("utf-8", errors="ignore")
                if result.stderr
                else "خطای نامشخص"
            )
            print(f"❌ خطا در نصب: {error_msg}")
            return False

    except subprocess.TimeoutExpired:
        print("⏰ زمان نصب به پایان رسید")
        return False
    except Exception as e:
        print(f"❌ خطای ناگهانی: {e}")
        return False


def quick_update_check():
    """
    بررسی سریع برای وجود نسخه جدید Edge و اعلام نیاز به بروزرسانی.

    Returns:
        bool or None: True اگر به‌روز است، False اگر نیاز به بروزرسانی دارد، None در صورت خطا.
    """
    try:
        current = get_current_edge_version()
        print(f"نسخه فعلی Edge: {current if current else 'نامشخص'}")

        latest_info = get_latest_edge_version()
        if latest_info and latest_info["version"]:
            print(f"آخرین نسخه: {latest_info['version']}")

            if current and current != latest_info["version"]:
                print("🔄 بروزرسانی موجود است!")
                return False  # نیاز به بروزرسانی
            else:
                print("✅ Edge به‌روز است")
                return True  # به‌روز است
        else:
            print("⚠️ نمی‌توان آخرین نسخه را بررسی کرد")
            return None

    except Exception as e:
        print(f"❌ خطا در بررسی: {e}")
        return None


def download_edge_driver_with_version_check() -> Optional[str]:
    """
    EdgeDriver را متناسب با نسخه مرورگر Edge دانلود و اکسترکت می‌کند.

    Returns:
        str or None: مسیر فایل msedgedriver.exe یا None در صورت خطا.
    """

    # 1. ایجاد پوشه webdrivers در کنار اسکریپت
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    driver_dir = os.path.join(script_dir, "webdrivers")
    os.makedirs(driver_dir, exist_ok=True)

    def get_edge_version() -> Optional[str]:
        """دریافت نسخه Edge"""
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Edge\BLBeacon"
            )
            version = winreg.QueryValueEx(key, "version")[0]
            winreg.CloseKey(key)
            return version
        except:
            return None

    def download_with_fallback(version: str) -> bool:
        """تلاش برای دانلود با نسخه مشخص و fallback"""
        urls = [
            f"https://msedgedriver.microsoft.com/{version}/edgedriver_win64.zip",
            f"https://msedgedriver.azureedge.net/{version}/edgedriver_win64.zip",
        ]

        try:
            main_version = version.split(".")[0]
            for i in range(5):
                fallback_version = str(int(main_version) - i)
                urls.extend(
                    [
                        f"https://msedgedriver.microsoft.com/{fallback_version}/edgedriver_win64.zip",
                        f"https://msedgedriver.azureedge.net/{fallback_version}/edgedriver_win64.zip",
                    ]
                )
        except:
            pass

        urls.extend(
            [
                "https://msedgedriver.azureedge.net/LATEST_STABLE",
            ]
        )

        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

        for url in urls:
            try:
                print(f"تلاش برای دانلود از: {url}")

                if "LATEST_STABLE" in url:
                    response = requests.get(url, headers=headers, timeout=10)
                    if response.status_code == 200:
                        latest_version = response.text.strip()
                        url = f"https://msedgedriver.azureedge.net/{latest_version}/edgedriver_win64.zip"
                        print(f"دانلود آخرین نسخه: {latest_version}")
                        response = requests.get(url, headers=headers, timeout=30)
                else:
                    response = requests.get(url, headers=headers, timeout=30)

                if response.status_code == 200:
                    # ذخیره در پوشه webdrivers
                    zip_path = os.path.join(driver_dir, "msedgedriver.zip")
                    with open(zip_path, "wb") as f:
                        f.write(response.content)

                    # استخراج به پوشه webdrivers
                    with zipfile.ZipFile(zip_path, "r") as zip_ref:
                        zip_ref.extractall(driver_dir)

                    os.remove(zip_path)
                    print("✅ Edge Driver با موفقیت دانلود و اکسترکت شد!")
                    return True

            except Exception as e:
                print(f"خطا در دانلود: {str(e)}")
                continue

        return False

    # اجرای اصلی
    print("🔍 در حال بررسی نسخه Edge...")
    edge_version = get_edge_version()

    if edge_version:
        print(f"نسخه Edge: {edge_version}")
    else:
        print("نسخه Edge پیدا نشد")

    version_to_use = edge_version if edge_version else "LATEST_STABLE"
    success = download_with_fallback(version_to_use)

    if not success:
        print("❌ دانلود ناموفق بود")
        return None

    driver_path = os.path.join(driver_dir, "msedgedriver.exe")
    if os.path.exists(driver_path):
        return driver_path
    else:
        print("❌ فایل msedgedriver.exe پیدا نشد")
        return None


def create_driver(
    headless: bool = False,
    ua: str | None = None,
    driver_path: str | None = None,  # مسیر اختیاری
) -> webdriver.Edge:
    """
    یک شیء WebDriver برای مرورگر Edge با تنظیمات دلخواه ایجاد می‌کند.

    Args:
        headless (bool): اجرای headless (بدون رابط گرافیکی).
        ua (str, optional): User-Agent سفارشی.
        driver_path (str, optional): مسیر EdgeDriver.

    Returns:
        webdriver.Edge: شیء WebDriver آماده استفاده.
    """
    # 1. پایان فرآیندهای مرورگر
    kill_process("msedge.exe")

    # 2. تعیین مسیر درایور
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    default_driver_path = os.path.join(script_dir, "webdrivers", "msedgedriver.exe")

    final_driver_path = driver_path or default_driver_path

    # 3. دانلود در صورت عدم وجود درایور
    if not os.path.exists(final_driver_path):
        print("⚠️ درایور یافت نشد. شروع دانلود...")
        downloaded_path = download_edge_driver_with_version_check()
        if downloaded_path:
            final_driver_path = downloaded_path
        else:
            raise FileNotFoundError("دانلود درایور ناموفق بود")

    # 4. تنظیمات Edge
    opts = webdriver.EdgeOptions()
    opts.add_argument("--inprivate")
    opts.add_argument(f"--remote-debugging-port={random.randint(9000, 9999)}")

    if headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")

    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    if not ua:
        ua = (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/125.0.0.0 Safari/537.36 Edg/125.0.0.0"
        )
    opts.add_argument(f"--user-agent={ua}")

    # 5. ایجاد سرویس و درایور
    try:
        service = Service(executable_path=final_driver_path)
        driver = webdriver.Edge(service=service, options=opts)

        # 6. ضد تشخیص
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {
                "source": """
                Object.defineProperty(navigator, 'webdriver', { 
                    get: () => undefined 
                });
                window.chrome = { runtime: {} };
                """
            },
        )

        return driver

    except Exception as e:
        print(f"خطا در ایجاد درایور: {str(e)}")
        raise


def wait_for_element(
    driver: webdriver.Edge,
    by: By,
    value: str,
    timeout: int = 10,
    visible: bool = False,
    clickable: bool = False,
):
    """
    منتظر ظاهر شدن یک عنصر در صفحه با شرط دلخواه می‌ماند.

    Args:
        driver (webdriver.Edge): شیء WebDriver.
        by (By): نوع جستجو (مثلاً By.ID).
        value (str): مقدار سلکتور.
        timeout (int): حداکثر زمان انتظار (ثانیه).
        visible (bool): اگر True، تا زمانی که عنصر قابل مشاهده شود صبر می‌کند.
        clickable (bool): اگر True، تا زمانی که عنصر قابل کلیک شود صبر می‌کند.

    Returns:
        WebElement or None: عنصر پیدا شده یا None در صورت خطا.
    """
    try:
        if clickable:
            condition = EC.element_to_be_clickable((by, value))
        elif visible:
            condition = EC.visibility_of_element_located((by, value))
        else:
            condition = EC.presence_of_element_located((by, value))

        element = WebDriverWait(driver, timeout).until(condition)
        return element

    except Exception as e:
        print(f"Error waiting for element ({value}): {e}")
        return None


def wait_for_elements(
    driver: webdriver.Edge,
    by: By,
    value: str,
    timeout: int = 10,
    visible: bool = False,
    clickable: bool = False,
) -> Optional[List[webdriver.remote.webelement.WebElement]]:
    """
    منتظر ظاهر شدن یک یا چند عنصر در صفحه با شرط دلخواه می‌ماند.

    Args:
        driver (webdriver.Edge): شیء WebDriver.
        by (By): نوع جستجو (مثلاً By.CSS_SELECTOR).
        value (str): مقدار سلکتور.
        timeout (int): حداکثر زمان انتظار (ثانیه).
        visible (bool): اگر True، تا زمانی که همه عناصر قابل مشاهده شوند صبر می‌کند.
        clickable (bool): اگر True، تا اولین عنصر قابل کلیک صبر می‌کند و سپس همه عناصر را برمی‌گرداند.

    Returns:
        list[WebElement] or None: لیست عناصر یا None در صورت خطا.
    """
    try:
        if clickable:
            # صبر می‌کند تا *حداقل یک* عنصر کلیک‌پذیر شود
            WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((by, value))
            )
            # سپس همه عناصری که الان حاضرند را می‌گیرد
            elements = driver.find_elements(by, value)

        elif visible:
            elements = WebDriverWait(driver, timeout).until(
                EC.visibility_of_all_elements_located((by, value))
            )
        else:
            elements = WebDriverWait(driver, timeout).until(
                EC.presence_of_all_elements_located((by, value))
            )

        return elements

    except Exception as e:
        print(f"Error waiting for elements ({value}): {e}")
        return None


def human_type(
    element: WebElement, text: str, min_delay: float = 0.01, max_delay: float = 0.1
):
    """
    متن را به صورت شبیه‌سازی تایپ انسانی در یک عنصر ورودی وارد می‌کند.

    Args:
        element (WebElement): عنصر ورودی.
        text (str): متن مورد نظر برای تایپ.
        min_delay (float): حداقل تاخیر بین کاراکترها (ثانیه).
        max_delay (float): حداکثر تاخیر بین کاراکترها (ثانیه).
    """
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(min_delay, max_delay))


def human_sleep(min_seconds: float = 2.0, max_seconds: float = 4.0):
    """
    یک مکث تصادفی شبیه انسان بین دو عمل ایجاد می‌کند.

    Args:
        min_seconds (float): حداقل مدت مکث (ثانیه).
        max_seconds (float): حداکثر مدت مکث (ثانیه).
    """
    time.sleep(random.uniform(min_seconds, max_seconds))


def scroll_to_bottom(
    driver: webdriver.Edge, step: int = 350, delay_range: tuple = (0.2, 0.3)
):
    """
    صفحه را به صورت تدریجی و شبیه انسان تا انتها اسکرول می‌کند.

    Args:
        driver (webdriver.Edge): شیء WebDriver.
        step (int): تعداد پیکسل برای هر بار اسکرول.
        delay_range (tuple): بازه تاخیر تصادفی بین اسکرول‌ها (ثانیه).
    """
    last_height = driver.execute_script("return document.body.scrollHeight")

    current_position = 0
    while current_position < last_height:
        driver.execute_script(f"window.scrollTo(0, {current_position});")
        current_position += step
        time.sleep(random.uniform(*delay_range))
        last_height = driver.execute_script("return document.body.scrollHeight")

    # در انتها یه مکث کوچک برای بارگذاری نهایی
    human_sleep(2, 4)
