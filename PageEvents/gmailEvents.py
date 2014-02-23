from bs4 import BeautifulSoup
import re
from xlrd import open_workbook
from xlwt import easyxf, Formula, Workbook
from dateutil.parser import parse
from dateutil.tz import tzlocal
import config


class orderDetails:
    orderDate = None
    Particulars = None
    Total_Amount = None
    Revenue_Amount = None


def extract_orders_from_email(mail, Con):
    try:
        html = mail.html
        OrderDate = mail.headers['Date']
        OrderDate = parse(OrderDate)
        OrderDate = OrderDate.astimezone(tzlocal())
        OrderDate = OrderDate.date()
        if OrderDate < Con.FromDate or OrderDate > Con.ToDate:
            return
        Particulars = get_particulars_from_html(mail.html)
        Total_Amount = get_total_from_html(mail.html)
        curr_order = orderDetails()
        curr_order.orderDate = str(OrderDate)
        curr_order.Particulars = Particulars
        curr_order.Total_Amount = Total_Amount
        curr_order.Revenue_Amount = round(float(Total_Amount) * 4.0 / 5.0, 2)
        Con.Orders.append(curr_order)
    except AttributeError:
        return


def get_particulars_from_html(html):
    soup = BeautifulSoup(html, 'html5lib')
    all_tables = soup.find_all('table')
    final_table = None
    for table in all_tables:
        if len(table.find_all('table')) > 0:
            continue
        elif 'ITEM' in table.text and 'QTY' in table.text and 'PAID' in table.text:
            final_table = table
            break
        else:
            continue

    trs = final_table.find_all('tr')
    Particulars = None
    for tr in trs:
        if 'ITEM' in tr.text and 'QTY' in tr.text and 'PAID' in tr.text:
            continue
        elif Particulars:
            Particulars += '\n' + tr.find_all('td')[0].text.strip()
        else:
            Particulars = tr.find_all('td')[0].text.strip()

    return Particulars


def get_total_from_html(html):
    soup = BeautifulSoup(html, 'html5lib')
    div_total = soup.find('div', text=re.compile('TOTAL:'))
    div_total = div_total.text.strip()
    div_total = div_total.replace('TOTAL:', '').strip()
    div_total = div_total.replace('$', '').strip()
    return div_total


def calculate_total_amount(Con):
    amount = 0.0
    revenue_amount = 0.0
    for order in Con.Orders:
        amount += float(order.Total_Amount)
        revenue_amount += float(order.Revenue_Amount)

    Con.Total_Amount = round(amount, 2)
    Con.Revenue_Amount = round(revenue_amount, 2)


def generate_xls(Con):
    wb = Workbook()
    ws = wb.add_sheet(sheetname='Order Results')
    headerStyle = easyxf('font: bold True')
    wrapStyle = easyxf('alignment: wrap True')
    ws.write(0, 0, 'Date', headerStyle)
    ws.write(0, 1, 'Particulars', headerStyle)
    ws.write(0, 2, 'Amount', headerStyle)
    ws.write(0, 3, 'Revenue', headerStyle)
    max_width_col0 = 0
    max_width_col1 = 0
    i = 1
    for order in Con.Orders:
        ws.write(i, 0, order.orderDate)
        ws.write(i, 1, order.Particulars, wrapStyle)
        ws.write(i, 2, float(order.Total_Amount))
        ws.write(i, 3, float(order.Revenue_Amount))
        for s in order.Particulars.split('\n'):
            if max_width_col1 < len(s):
                max_width_col1 = len(s)

        if max_width_col0 < len(str(order.orderDate)):
            max_width_col0 = len(str(order.orderDate))
        i += 1

    ws.col(1).width = 256 * max_width_col1 + 2
    ws.col(0).width = 256 * max_width_col0 + 20
    ws.write(i, 1, 'Total Amount', headerStyle)
    ws.write(i, 2, Con.Total_Amount, headerStyle)
    ws.write(i, 3, Con.Revenue_Amount, headerStyle)
    filename = Con.Username + '_' + Con.Label + '_'
    filename += Con.FromDate.strftime('%Y%m%d') + '_'
    filename += Con.ToDate.strftime('%Y%m%d') + '.xls'
    wb.save(config.get_main_dir() + '/Output/' + filename)