import requests
from bs4 import BeautifulSoup
import time
import re
from sheets_auth import get_sheets_service, extract_spreadsheet_id, write_to_sheet, read_from_sheet
from chrome_driver_setup import get_chrome_driver, quit_driver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# OAuth 2.0クライアントの設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def scrape_yahoo_auction(url):
    driver = None
    try:
        print(f"Scraping URL: {url}")
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        try:
            title = soup.select_one('div.ProductTitle__title h1').text.strip()
            print(f"Title: {title}")
        except AttributeError:
            print("Error: Could not find title")
            title = 'N/A'

        try:
            auction_id = soup.select_one('th:-soup-contains("オークションID") + td').text.strip()
            print(f"Auction ID: {auction_id}")
        except AttributeError:
            print("Error: Could not find auction ID")
            auction_id = 'N/A'

        seller_link = soup.select_one('a[href^="https://auctions.yahoo.co.jp/seller/"][data-cl-params*="seller"]')
        if seller_link and 'href' in seller_link.attrs:
            seller_url = seller_link['href']
            seller_id = seller_url.split('/seller/')[-1]
            seller = seller_link.text.strip()
            print(f"Seller ID: {seller_id}, Seller Name: {seller}")
        else:
            print("Error: Could not find seller information")
            seller_id = 'N/A'
            seller = 'N/A'

        try:
            end_date = soup.select_one('th:-soup-contains("終了日時") + td').text.strip()
            print(f"End Date: {end_date}")
        except AttributeError:
            print("Error: Could not find end date")
            end_date = 'N/A'

        price_element = soup.select_one('dd.Price__value')
        if price_element:
            price_text = price_element.contents[0].strip()
            price = re.sub(r'[^\d]', '', price_text)
            print(f"Price: {price}")
        else:
            print("Error: Could not find price")
            price = 'N/A'

        try:
            tax_included_price = soup.select_one('span.Price__tax').text.strip()
            tax_included_price = re.sub(r'[^\d]', '', tax_included_price)
            print(f"Tax Included Price: {tax_included_price}")
        except AttributeError:
            print("Error: Could not find tax included price")
            tax_included_price = '0'

        # tax_included_priceが'0'の場合、priceの値を使用
        if tax_included_price == '0':
            tax_included_price = price
            print(f"Using price as tax included price: {tax_included_price}")

        # 送料情報の抽出（Seleniumを使用）
        driver = get_chrome_driver()
        driver.get(url)
        try:
            # まず、JavaScriptが読み込まれるのを待つ
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
            )

            # 送料情報を取得するための複数のセレクタを試す
            selectors = [
                "span.PricepostageValue",
                "span.Price__postageValue",
                "dd.Price__postage",
                "span[data-react-unit-name='PostageValue']"
            ]

            postage_element = None
            for selector in selectors:
                try:
                    postage_element = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    if postage_element:
                        break
                except TimeoutException:
                    continue

            if postage_element:
                postage_text = postage_element.text.strip()
                print(f"Raw postage text (Selenium): {postage_text}")

                postage_match = re.search(r'([\d,]+)円', postage_text)
                if postage_match:
                    total_postage = postage_match.group(1).replace(',', '')
                    print(f"Extracted total postage: {total_postage}")
                elif '着払い' in postage_text or '落札者負担' in postage_text:
                    total_postage = '着払い'
                    print("Total postage is buyer's responsibility or cash on delivery")
                else:
                    total_postage = postage_text
                    print(f"Unrecognized total postage format: {postage_text}")
            else:
                print("Error: Could not find postage element")
                total_postage = ''
        except TimeoutException:
            print("Error: Timed out waiting for postage element")
            total_postage = ''
        except NoSuchElementException:
            print("Error: Postage element not found")
            total_postage = ''
        except Exception as e:
            print(f"Error extracting total postage with Selenium: {str(e)}")
            total_postage = ''

        return {
            'title': title,
            'transaction_id': auction_id,
            'seller_id': seller_id,
            'seller_name': seller,
            'transaction_date': end_date,
            'price': price,
            'tax_included_price': tax_included_price,
            'total_postage': total_postage
        }

    except Exception as e:
        print(f"Error processing URL: {url}. Error: {str(e)}")
        return None
    finally:
        if driver:
            quit_driver(driver)

def smart_scraping(service, spreadsheet_url, sheet_name, start_row, end_row, is_scraping):
    try:
        print(f"Starting smart_scraping for sheet: {sheet_name}, rows: {start_row} to {end_row}")

        spreadsheet_id = extract_spreadsheet_id(spreadsheet_url)
        if not spreadsheet_id:
            raise ValueError("Invalid spreadsheet URL")
        print(f"Extracted spreadsheet ID: {spreadsheet_id}")

        # タグ行（2行目）を取得
        try:
            tags = read_from_sheet(service, spreadsheet_url, f'{sheet_name}!2:2')[0]
            print(f"Retrieved tags: {tags}")
        except Exception as e:
            print(f"Error retrieving tags: {str(e)}")
            raise

        # 'url' タグの列インデックスを取得
        try:
            url_index = tags.index('url')
            print(f"URL index: {url_index}")
        except ValueError:
            print("Error: 'url' tag not found in the second row.")
            raise ValueError("'url'タグが2行目に見つかりません。")

        # 処理する行の範囲を決定
        range_to_process = f'{sheet_name}!{start_row}:{end_row if end_row else ""}'
        try:
            rows = read_from_sheet(service, spreadsheet_url, range_to_process)
            print(f"Retrieved {len(rows)} rows to process")
        except Exception as e:
            print(f"Error reading rows from sheet: {str(e)}")
            raise

        new_data_count = 0
        skipped_count = 0

        for row_num, row in enumerate(rows, start=start_row):
            if not is_scraping():
                print("スクレイピングが中断されました。")
                break

            print(f"\nProcessing row {row_num}")
            # 行のデータが足りない場合、空文字で埋める
            row_data = row + [''] * (len(tags) - len(row))

            if len(row_data) <= url_index or not row_data[url_index].strip():
                print(f"行 {row_num}: URLが空です。スキップします。")
                skipped_count += 1
                continue

            url = row_data[url_index].strip()
            print(f"行 {row_num}: URL {url} の処理を開始します。")

            # スクレイピング処理
            scraped_data = scrape_yahoo_auction(url)

            if scraped_data:
                print(f"行 {row_num}: スクレイピング成功")
                print(f"Scraped data: {scraped_data}")

                # 更新が必要な列だけを特定してバッチ更新用のデータを準備
                updates = []
                for tag, value in scraped_data.items():
                    if tag in tags:
                        col_letter = chr(65 + tags.index(tag))  # A, B, C...の列文字を生成
                        range_name = f'{sheet_name}!{col_letter}{row_num}'
                        updates.append({
                            'range': range_name,
                            'values': [[value]]
                        })

                # バッチ更新を実行
                if updates:
                    try:
                        body = {
                            'valueInputOption': 'USER_ENTERED',
                            'data': updates
                        }
                        service.spreadsheets().values().batchUpdate(
                            spreadsheetId=spreadsheet_id,
                            body=body
                        ).execute()
                        print(f"Successfully updated row {row_num} with specific columns")
                        print(f"Updated ranges: {[update['range'] for update in updates]}")
                        new_data_count += 1
                    except Exception as e:
                        print(f"Error updating row {row_num}: {str(e)}")
                        continue

            else:
                print(f"行 {row_num}: スクレイピング失敗")

            # レート制限（1秒待機）
            time.sleep(1)

        print(f"\nScraping completed. New data count: {new_data_count}, Skipped count: {skipped_count}")
        return new_data_count, skipped_count

    except Exception as e:
        print(f"Error in smart_scraping: {str(e)}")
        raise

def main(spreadsheet_url, start_row, end_row, sheet_name, is_scraping):
    try:
        service = get_sheets_service()

        new_count, skipped_count = smart_scraping(service, spreadsheet_url, sheet_name, start_row, end_row, is_scraping)

        return new_count, skipped_count

    except ValueError as e:
        print(f"エラー: {str(e)}")
        raise
    except Exception as e:
        print(f"予期せぬエラーが発生しました: {str(e)}")
        raise

if __name__ == "__main__":
    # テスト用のURL（実際の使用時はGUIから渡される）
    test_url = "https://docs.google.com/spreadsheets/d/your-spreadsheet-id/edit#gid=0"
    new_count, skipped_count = main(test_url, 3, None, "ヤフオク購入履歴", lambda: True)
    print(f"新たにスクレイピングしたURL数: {new_count}")
    print(f"スキップしたURL数: {skipped_count}")