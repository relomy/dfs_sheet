import requests
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Font, colors
# from openpyxl.cell import get_column_letter
from os import path


def print_header(sheet):
    """ print header to sheet """
    # set up columns and column widths
    columns = ["IP", "Date (UTC)", "Method", "URL", "Referrer", "User Agent"]
    column_widths = [18, 20, 8, 80, 25, 100]

    for col_num, header_field in enumerate(columns):
        sheet.cell(row=1, column=col_num + 1).value = header_field
    light_orange = "00FFB732"
    for row in sheet.iter_rows('A1:H1'):
        for cell in row:
            cell.font = Font(bold=True, color=colors.BLACK)
            cell.border = Border(bottom=Side(style="thin"))
            cell.fill = PatternFill(patternType='solid', start_color=light_orange, end_color=light_orange)


def get_nfl_receptions():
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/receptions/2018/RB'
    # set parameters
    # params = {
    #     'leagueId': league_id,
    #     'seasonId': year
    # }

    # send GET request
    r = requests.get(ENDPOINT)
    status = r.status_code

    # if not successful, raise an exception
    if status != 200:
        raise Exception('[script.py] Requests status != 200. It is: {0}'.format(status))

    # store response
    data = r.json()


def print_position_ws(wb, position, fields):
    if position in wb.sheetnames:
        wb[position].append(fields)
        # wsx = wb.create_sheet(title=position)
    else:
        # if wb[position] does not exist, create it and print header
        wb.create_sheet(title=position)
        header = ['Position', 'Name', 'Salary', 'TeamAbbrev', 'AvgPointsPerGame']
        wb[position].append(header)
        wb[position].append(fields)


def main():
    get_nfl_receptions()
    exit()
    fn = 'DKSalaries.csv'
    dest_filename = 'sheet.xlsx'

    # create workbook/worksheet
    wb = Workbook()
    ws1 = wb.active
    ws1.title = 'DKSalaries'

    # guess types (numbers, floats, etc)
    wb.guess_types = True

    # ws2 = wb.create_sheet(title='QB')

    with open(fn, 'r') as f:
        # read entire file into memory
        lines = f.readlines()

        for i, line in enumerate(lines):
            # append header to first worksheet, otherwise skip it
            if i == 0:
                ws1.append(line.rstrip().split(','))
                continue

            fields = line.rstrip().split(',')
            my_fields = [fields[0], fields[2], fields[5], fields[7], fields[8]]

            # if i == 0:
                # header - columns 0, 2, 5, 7, 8
                # position, name, salary, teamabbv, avgppg
                # ws2.append(header)

            # print fields to position-named worksheet
            print_position_ws(wb, fields[0], my_fields)

            # print fields to first worksheet
            ws1.append(line.rstrip().split(','))

        # for i, line in enumerate(lines):
        #     fields = line.rstrip().split(',')
        #
        # for line in f:
        #     for field in line.split(','):
        #         ws1.cell()

    # save workbook (.xlsx file)
    wb.save(filename=dest_filename)


if __name__ == "__main__":
    main()
