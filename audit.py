import asyncio
from playwright.async_api import async_playwright
from datetime import datetime, timedelta
import time
import os
import shutil
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import re

DOWNLOAD_DIR = "/tmp"

def rename_downloaded_file(download_dir, download_path, prefix):
    try:
        current_hour = datetime.now().strftime("%H")
        new_file_name = f"{prefix}-{current_hour}.csv"
        new_file_path = os.path.join(download_dir, new_file_name)
        if os.path.exists(new_file_path):
            os.remove(new_file_path)
        shutil.move(download_path, new_file_path)
        print(f"Arquivo salvo como: {new_file_path}")
        return new_file_path
    except Exception as e:
        print(f"Erro ao renomear o arquivo: {e}")
        return None

def update_google_sheets(csv_file_path, aba_nome):
    try:
        if not os.path.exists(csv_file_path):
            print(f"Arquivo {csv_file_path} não encontrado.")
            return
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("hxh.json", scope)
        client = gspread.authorize(creds)
        sheet1 = client.open_by_url("https://docs.google.com/spreadsheets/d/1ZCpHclzg4aiaco5Nk1dWX4Q3zv6II3p0wfFDUkTB7pg/edit?gid=0#gid=0")
        worksheet1 = sheet1.worksheet(aba_nome)
        df = pd.read_csv(csv_file_path)
        df = df.fillna("")
        worksheet1.clear()
        worksheet1.update([df.columns.values.tolist()] + df.values.tolist())
        print(f"Arquivo enviado com sucesso para a aba '{aba_nome}'.")
        time.sleep(5)
    except Exception as e:
        print(f"Erro durante o processo: {e}")

async def main():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, args=["--no-sandbox", "--disable-dev-shm-usage", "--window-size=1920,1080"])
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()
        try:
            # LOGIN
            await page.goto("https://spx.shopee.com.br/")
            await page.wait_for_selector('xpath=//*[@placeholder="Ops ID"]', timeout=15000)
            await page.locator('xpath=//*[@placeholder="Ops ID"]').fill('Ops92710')
            await page.locator('xpath=//*[@placeholder="Senha"]').fill('@Shopee123')
            await page.locator('xpath=/html/body/div[1]/div/div[2]/div/div/div[1]/div[3]/form/div/div/button').click()
            await page.wait_for_timeout(15000)
            try:
                await page.locator('.ssc-dialog-close').click(timeout=5000)
            except:
                print("Nenhum pop-up foi encontrado.")
                await page.keyboard.press("Escape")

            """
            # DOWNLOAD PS
            await page.goto("https://spx.shopee.com.br/#/dashboard/toProductivity?page_type=Outbound")
            await page.wait_for_timeout(10000)
            await page.locator('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div[2]/div[3]/span/span/span/button').click()
            await page.wait_for_timeout(10000)
            await page.locator("div").filter(has_text=re.compile("^Exportar$")).click()
            await page.wait_for_timeout(10000)
            async with page.expect_download() as download_info:
                await page.locator('xpath=/html/body/span/div/div[1]/div/span/div/div[2]/div[2]/div[1]/div/div[1]/div/div[1]/div[2]/button').click()
            download_prod = await download_info.value
            prod_download_path = os.path.join(DOWNLOAD_DIR, download_prod.suggested_filename)
            await download_prod.save_as(prod_download_path)
            prod_file_path = rename_downloaded_file(DOWNLOAD_DIR, prod_download_path, "PS")
            """

            # DOWNLOAD AUD
            await page.goto("https://spx.shopee.com.br/#/audit-management/to-audit")
            await page.wait_for_timeout(10000)
            await page.locator('xpath=/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/span[1]/span[1]/button[1]/span[1]').click()
            await page.wait_for_timeout(10000)
            #await page.locator('xpath=').click()
            await page.get_by_role("menuitem", name="Exportar").nth(0).click()
            #d1 = (datetime.now() - timedelta(days=1)).strftime("%Y/%m/%d")
            #date_input = page.get_by_role("textbox", name="Escolha a data de início").nth(0)
            #await date_input.wait_for(state="visible", timeout=10000)
            #await date_input.click(force=True)
            #await date_input.fill(d1)
            await page.wait_for_timeout(5000)
            #await page.locator('xpath=/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[6]/form[1]/div[4]/button[1]').click()
            #await page.locator('xpath=/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[8]/div[1]/button[1]').click()
            await page.wait_for_timeout(10000)
            async with page.expect_download() as download_info:
                await page.locator('xpath=/html/body/span/div/div[1]/div/span/div/div[2]/div[2]/div[1]/div/div[1]/div/div[1]/div[2]/button').click()
            download_ws = await download_info.value
            ws_download_path = os.path.join(DOWNLOAD_DIR, download_ws.suggested_filename)
            await download_ws.save_as(ws_download_path)
            ws_file_path = rename_downloaded_file(DOWNLOAD_DIR, ws_download_path, "AUD")

            # Atualizar Google Sheets
            #if prod_file_path:
            #    update_google_sheets(prod_file_path, "PS")
            if ws_file_path:
                update_google_sheets(ws_file_path, "AUDITORIA")
            print("Dados atualizados com sucesso.")
        except Exception as e:
            print(f"Erro durante o processo: {e}")
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
