import time

import browser_automation as br
import excel_operations as ex
import email_operations as em
import date_operations as da

WORKDIR = r'S:\Work\Internship\GreenAtom\RPAmoex'
SMTP_SERVER = 'smtp.mail.ru'
SMTP_PORT = 587
SMTP_USER = 'b.r.o.l@mail.ru'
SMTP_PASSWORD = 'bn3iwJxd6igUeYQJjQa3'
IMAP_SERVER = 'imap.mail.ru'
IMAP_USER = 'b.r.o.l@mail.ru'
IMAP_PASSWORD = 'bn3iwJxd6igUeYQJjQa3'

try:
    driver = br.open_website()

    ex.initialize_excel_file(f"{WORKDIR}\\excelResults.xlsx")

    previous_month, last_day_pm, first_weekday_pm = da.get_previous_month_details()

    xpath_from_PM, xpath_to_PM, xpath_first_day, xpath_last_day = br.generate_xpaths(previous_month, first_weekday_pm,
                                                                                     last_day_pm)

    br.navigate_to_currency_pairs(driver)

    # USD-RUB
    br.setup_currency_pair(driver, xpath_from_PM, xpath_to_PM, xpath_first_day, xpath_last_day,
                           '//div[18]//a', "USD-RUB")

    usd_data = br.extract_usd_data(driver)

    ex.write_usd_data_to_excel(usd_data, f"{WORKDIR}\\excelResults.xlsx")

    # JPY-RUB
    br.setup_currency_pair(driver, xpath_from_PM, xpath_to_PM, xpath_first_day, xpath_last_day,
                           '//div[3]/div[1]/div[8]', "JPY_RUB")

    jpy_data = br.extract_jpy_data(driver)

    ex.write_jpy_data_to_excel(jpy_data, f"{WORKDIR}\\excelResults.xlsx")

    ex.calculate_usd_to_jpy_ratio(f"{WORKDIR}\\excelResults.xlsx")

    ex.format_excel_file(f"{WORKDIR}\\excelResults.xlsx")

    em.send_email_with_report(SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, IMAP_SERVER, IMAP_USER, IMAP_PASSWORD,
                              f"{WORKDIR}\\excelResults.xlsx")
except Exception as e:
    print(e)
finally:
    time.sleep(3)
    driver.quit()
