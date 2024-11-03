import os
import requests
import zipfile
import io
import platform
import logging
import json
import time
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

CHROME_DRIVER_FOLDER = os.path.join(os.path.expanduser("~"), ".chrome_driver")
CHROME_DRIVER_PATH = os.path.join(CHROME_DRIVER_FOLDER, "chromedriver")
JSON_ENDPOINT = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"


def get_platform():
    system = platform.system()
    if system == "Darwin":  # macOS
        return "mac-arm64" if platform.machine() == "arm64" else "mac-x64"
    elif system == "Windows":
        return "win64" if platform.machine().endswith('64') else "win32"
    elif system == "Linux":
        return "linux64"
    else:
        raise OSError("Unsupported operating system")


def get_chrome_version():
    if platform.system() == "Darwin":  # macOS
        cmd = r'/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --version'
    elif platform.system() == "Windows":
        cmd = r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version'
    else:
        raise OSError("Unsupported operating system")

    output = os.popen(cmd).read()
    version = output.strip().split()[-1]
    logging.info(f"Detected Chrome version: {version}")
    return version


def get_matching_driver_version(chrome_version):
    try:
        response = requests.get(JSON_ENDPOINT)
        response.raise_for_status()
        data = response.json()
        chrome_major_minor = '.'.join(chrome_version.split('.')[:2])  # 例: "128.0"
        compatible_versions = []
        for version in data['versions']:
            if version['version'].startswith(chrome_major_minor):
                compatible_versions.append(version['version'])
        if compatible_versions:
            selected_version = max(compatible_versions, key=lambda x: [int(i) for i in x.split('.')])
            logging.info(f"Selected ChromeDriver version: {selected_version} for Chrome {chrome_version}")
            return selected_version
        logging.error(f"No matching ChromeDriver version found for Chrome {chrome_version}")
        return None
    except requests.RequestException as e:
        logging.error(f"Error fetching ChromeDriver version: {e}")
        return None


def get_download_url_from_json(version):
    try:
        response = requests.get(JSON_ENDPOINT)
        response.raise_for_status()
        data = response.json()
        platform_name = get_platform()
        logging.info(f"Searching for ChromeDriver version {version} for platform {platform_name}")
        for v in data['versions']:
            if v['version'] == version:
                for download in v['downloads']['chromedriver']:
                    if download['platform'] == platform_name:
                        logging.info(f"Found download URL: {download['url']}")
                        return download['url']
        logging.error(f"No matching download URL found for version {version} and platform {platform_name}")
        return None
    except requests.RequestException as e:
        logging.error(f"Error fetching download URL from JSON: {e}")
        return None
    except KeyError as e:
        logging.error(f"Unexpected JSON structure: {e}")
        return None


def download_driver(version):
    if not os.path.exists(CHROME_DRIVER_FOLDER):
        os.makedirs(CHROME_DRIVER_FOLDER)

    url = get_download_url_from_json(version)
    if not url:
        logging.error(f"Unable to find download URL for ChromeDriver version {version}")
        return False

    logging.info(f"Attempting to download ChromeDriver from: {url}")

    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()
        total_size = int(response.headers.get('content-length', 0))
        block_size = 1024  # 1 KB
        zip_path = os.path.join(CHROME_DRIVER_FOLDER, "chromedriver.zip")
        with open(zip_path, 'wb') as file, \
                tqdm(desc="Downloading ChromeDriver", total=total_size, unit='iB', unit_scale=True) as progress_bar:
            for data in response.iter_content(block_size):
                size = file.write(data)
                progress_bar.update(size)

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(CHROME_DRIVER_FOLDER)

        # リストZIPファイルの内容
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            logging.info(f"ZIP contents: {zip_ref.namelist()}")

        # ChromeDriverファイルを探して適切な場所に移動
        for root, dirs, files in os.walk(CHROME_DRIVER_FOLDER):
            for file in files:
                if file.startswith('chromedriver'):
                    os.rename(os.path.join(root, file), CHROME_DRIVER_PATH)
                    break

        if not os.path.exists(CHROME_DRIVER_PATH):
            raise FileNotFoundError(f"ChromeDriver not found in the extracted files")

        os.chmod(CHROME_DRIVER_PATH, 0o755)  # 実行権限を付与
        logging.info(f"Successfully downloaded and extracted ChromeDriver {version}")
        return True
    except Exception as e:
        logging.error(f"Unexpected error during ChromeDriver download and extraction: {e}")
    return False


def download_with_retry(version, max_retries=3, delay=5):
    for attempt in range(max_retries):
        if download_driver(version):
            return True
        logging.warning(f"Download attempt {attempt + 1} failed. Retrying in {delay} seconds...")
        time.sleep(delay)
    return False


def get_chrome_driver():
    try:
        chrome_version = get_chrome_version()
        driver_version = get_matching_driver_version(chrome_version)

        if not driver_version:
            raise Exception("Unable to determine appropriate ChromeDriver version")

        if not os.path.exists(CHROME_DRIVER_PATH) or get_driver_version() != driver_version:
            logging.info(f"Downloading ChromeDriver {driver_version}")
            if not download_with_retry(driver_version):
                raise Exception("Failed to download ChromeDriver after multiple attempts")

        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

        service = ChromeService(executable_path=CHROME_DRIVER_PATH)
        driver = webdriver.Chrome(service=service, options=chrome_options)

        logging.info("Chrome driver initialized successfully")
        return driver

    except Exception as e:
        logging.error(f"Error initializing Chrome driver: {e}")
        raise


def get_driver_version():
    try:
        output = os.popen(f"{CHROME_DRIVER_PATH} --version").read()
        return output.split()[1]
    except Exception as e:
        logging.error(f"Error getting ChromeDriver version: {e}")
        return None


def quit_driver(driver):
    try:
        if driver:
            driver.quit()
            logging.info("Chrome driver quit successfully")
    except Exception as e:
        logging.error(f"Error quitting Chrome driver: {e}")


if __name__ == "__main__":
    driver = get_chrome_driver()
    if driver:
        print("ChromeDriver initialized successfully")
        quit_driver(driver)
    else:
        print("Failed to initialize ChromeDriver")