import datetime
from sys import argv
from sqlalchemy import select, desc
from variables import WORKBOOK_DEFAULT, DB_DEFAULT, DATE1_DEFAULT, DATE2_DEFAULT
from xls_load import row_gen
from xls_db import XlsStorage

""" Main module for start XLSX_Parser """

def xls_parser(workdook=WORKBOOK_DEFAULT, db=DB_DEFAULT):
    current_data, date1, date2 = row_gen(workdook)
    xls_db = XlsStorage(db)
    company_names = []
    dates = []

    if date1 == DATE1_DEFAULT:
        date1 = datetime.datetime.now().date() + datetime.timedelta(days=-1)
        dates.append(date1)
    if date2 == DATE2_DEFAULT:
        date2 = datetime.datetime.now().date()

    dates = [date1, date2]

    for row in current_data:
        xls_id, company, fact_qliq_d1, fact_qliq_d2, fact_qoil_d1, fact_qoil_d2, \
            forecast_qliq_d1, forecast_qliq_d2, forecast_qoil_d1, forecast_qoil_d2 = list(row)
        line_dict = {
            'xls_id': xls_id,
            'company': company,
            'data': (
                [fact_qliq_d1, fact_qoil_d1, forecast_qliq_d1, forecast_qoil_d1, date1],
                [fact_qliq_d2, fact_qoil_d2, forecast_qliq_d2, forecast_qoil_d2, date2]
            )
        }

        if company not in company_names:
            company_names.append(company)

        for i in line_dict['data']:
            fact_qliq, fact_qoil, forecast_qliq, forecast_qliq, date = i
            row = select(xls_db.XlsTable).where(
                                    xls_db.XlsTable.date.is_(date),
                                    xls_db.XlsTable.company.is_(company)
                                    ).order_by(desc(xls_db.XlsTable.id)).limit(1)
            last_line_from_db = list(xls_db.session.scalars(row))

            if not last_line_from_db:
                fact_qliq_total = fact_qliq
                fact_qoil_total = fact_qoil
                xls_db.add_row(line_dict['xls_id'], line_dict['company'], fact_qliq, fact_qoil, forecast_qliq,
                               forecast_qliq, date, fact_qliq_total, fact_qoil_total)
            else:
                fact_qliq_total = last_line_from_db[0].total_qliq + fact_qliq
                fact_qoil_total = last_line_from_db[0].total_qoil + fact_qoil
                xls_db.add_row(xls_id, company, fact_qliq_d1, fact_qoil_d1, forecast_qliq, forecast_qliq, date,
                               fact_qliq_total, fact_qoil_total)

    last_lines_from_db = []
    for i_company in company_names:
        for i_date in dates:
            last_row= select(xls_db.XlsTable).where(xls_db.XlsTable.date.is_(i_date),
                                                    xls_db.XlsTable.company.in_([i_company])
                                                      ).order_by(desc(xls_db.XlsTable.id)).limit(1)
            last_line_from = list(xls_db.session.scalars(last_row))
            last_lines_from_db.append(last_line_from)

    return last_lines_from_db

if __name__ == '__main__':
    """ 
    Для запуска через коммандную строку:
        - Обычный запуск: python main.py
        - Запуск с параметрами: python main.py base_data.xlsx xls_db.db3
    """

    param = argv
    try:
        argv[1], argv[2]
        print(f'Параметры запуска: {param}')
        total = xls_parser(argv[1], argv[2])
    except:
        total = xls_parser(WORKBOOK_DEFAULT, DB_DEFAULT)
    total = xls_parser('base_data.xlsx')
    for i in total:
        print(f'Компания: {list(i)[0].company}, '
              f'Дата: {list(i)[0].date}, '
              f'Fact_QLiq_Total: {list(i)[0].total_qliq}, '
              f'Fact_QOil_Total:{list(i)[0].total_qoil}')
