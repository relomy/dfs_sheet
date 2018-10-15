import json
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font, colors
# from openpyxl.cell import get_column_letter
from os import path


def style_range(ws, cell_range, border=Border(), fill=None, font=None, alignment=None):
    """
    Apply styles to a range of cells as if they were a single cell.

    :param ws:  Excel worksheet instance
    :param range: An excel range to style (e.g. A1:F20)
    :param border: An openpyxl Border
    :param fill: An openpyxl PatternFill or GradientFill
    :param font: An openpyxl Font object
    """

    top = Border(top=border.top)
    left = Border(left=border.left)
    right = Border(right=border.right)
    bottom = Border(bottom=border.bottom)

    first_cell = ws[cell_range.split(":")[0]]
    if alignment:
        ws.merge_cells(cell_range)
        first_cell.alignment = alignment

    rows = ws[cell_range]
    if font:
        first_cell.font = font

    for cell in rows[0]:
        cell.border = cell.border + top
    for cell in rows[-1]:
        cell.border = cell.border + bottom

    for row in rows:
        l = row[0]
        r = row[-1]
        l.border = l.border + left
        r.border = r.border + right
        if fill:
            for c in row:
                c.fill = fill


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


def create_sheet_header(wb, title, header):
    wb.create_sheet(title=title)
    wb[title].append(header)


def pull_data(filename, ENDPOINT):
    """Either pull file from API or from file."""
    data = None
    if not path.isfile(filename):
        print("{} does not exist. Pulling from endpoint [{}]".format(filename, ENDPOINT))
        # send GET request
        r = requests.get(ENDPOINT)
        status = r.status_code

        # if not successful, raise an exception
        if status != 200:
            raise Exception('Requests status != 200. It is: {0}'.format(status))

        # store response
        data = r.json()

        # dump json to file for future use to avoid multiple API pulls
        with open(filename, 'w') as outfile:
            json.dump(data, outfile)
    else:
        print("File exists [{}]. Nice!".format(filename))
        # load json from file
        with open(filename, 'r') as json_file:
            data = json.load(json_file)

    return data


def get_nfl_snaps(wb):
    """Retrieve snaps from lineups.com API."""
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/snaps/2018/OFF'
    filename = 'nfl_snaps.json'

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = pull_data(filename, ENDPOINT)

    if data is None:
        raise Exception('Failed to pull data from API or file.')

    player_data = data['data']

    # create worksheet
    title = 'SNAPS'
    header = ['name', 'position', 'team', 'season average' 'week1', 'week2', 'week3', 'week4', 'week5', 'week6',
              'week7', 'week8', 'week9', 'week10', 'week11', 'week12', 'week13', 'week14',
              'week15', 'week16']
    create_sheet_header(wb, title, header)

    for d in player_data:
        name = d['full_name']
        position = d['position']
        team = d['team']
        weeks = d['snap_percentage_by_week']  # list
        season_average = d['season_snap_percent']

        # we only care about RB/TE/WR
        if position not in ['RB', 'TE', 'WR']:
            continue

        # convert weeks dict to list
        all_weeks = []
        for weekly_targets in weeks:
            # if weeks is None, put in blank string
            # 0 would mean they played but didn't get a snap
            if weekly_targets is None:
                all_weeks.append('')
            else:
                all_weeks.append(weekly_targets)

        # pad weeks to 16 (a = [])
        # more visual/pythonic
        # a = (a + N * [''])[:N]
        N = 16
        all_weeks = (all_weeks + N * [''])[:N]

        # add three lists together
        pre_weeks = [name, position, team, season_average]
        # post_weeks = [targets, average, recv_touchdowns]
        ls = pre_weeks + all_weeks

        wb[title].append(ls)


def get_nfl_targets(wb):
    """Retrieve targets from lineups.com API."""
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/targets/2018/OFF'
    filename = 'nfl_targets.json'

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = pull_data(filename, ENDPOINT)

    player_data = data['data']

    # create worksheet
    title = 'TARGETS'
    header = ['name', 'position', 'team', 'season average', 'week1', 'week2', 'week3', 'week4', 'week5', 'week6',
              'week7', 'week8', 'week9', 'week10', 'week11', 'week12', 'week13', 'week14',
              'week15', 'week16', 'targets', 'recv touchdowns']
    create_sheet_header(wb, title, header)

    for d in player_data:
        # TODO target percentage? it's by week as well
        name = d['full_name']
        position = d['position']
        team = d['team']
        targets = d['total']
        weeks = d['weeks']  # dict
        season_average = d['average']
        recv_touchdowns = d['receiving_touchdowns']
        catch_percentage = d['catch_percentage']
        season_target_percent = d['season_target_percent']

        # we only care about RB/TE/WR
        if position not in ['RB', 'TE', 'WR']:
            continue

        # convert weeks dict to list
        all_weeks = []
        for weekly_targets in weeks:
            # if weeks is None, put in blank string
            # 0 would mean they played but didn't get a snap
            if weekly_targets is None:
                all_weeks.append('')
            else:
                all_weeks.append(weekly_targets)

        # pad weeks to 16 (a = [])
        # more visual/pythonic
        # a = (a + N * [''])[:N]
        N = 16
        all_weeks = (all_weeks + N * [''])[:N]

        # add three lists together
        pre_weeks = [name, position, team, season_average]
        post_weeks = [targets, recv_touchdowns]
        ls = pre_weeks + all_weeks + post_weeks

        # insert all_weeks list into ls
        # ls = [name, position, rating, team, receptions, average, touchdowns]
        # print("trying to insert: ls[2:{}]".format(len(all_weeks)))
        # ls[4:len(all_weeks)-1] = all_weeks
        # print(ls)

        wb[title].append(ls)


def get_nfl_receptions(wb):
    """Retrieve receptions from lineups.com API."""
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/receptions/2018/OFF'
    filename = 'nfl_receptions.json'

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = pull_data(filename, ENDPOINT)

    # we just want player data
    player_data = data['data']

    # create worksheet
    title = 'RECEPTIONS'
    header = ['name', 'position', 'team', 'week1', 'week2', 'week3', 'week4', 'week5', 'week6',
              'week7', 'week8', 'week9', 'week10', 'week11', 'week12', 'week13', 'week14',
              'week15', 'week16', 'receptions', 'average', 'touchdowns']
    create_sheet_header(wb, title, header)

    for d in player_data:
        name = d['name']
        position = d['position']
        team = d['team']
        receptions = d['receptions']
        weeks = d['weeks']  # dict
        season_average = d['average']
        touchdowns = d['touchdowns']

        # we only care about RB/TE/WR
        if position not in ['RB', 'TE', 'WR']:
            continue

        # convert weeks dict to list
        all_weeks = []
        for i in range(0, len(weeks)):
            # if weeks is None, put in blank string
            # 0 would mean they played but didn't get a snap
            if weeks[str(i + 1)] is None:
                all_weeks.append('')
            else:
                all_weeks.append(weeks[str(i + 1)])

        # pad weeks to 16 (a = [])
        # more visual/pythonic
        # a = (a + N * [''])[:N]
        N = 16
        all_weeks = (all_weeks + N * [''])[:N]

        # add three lists together
        pre_weeks = [name, position, team, season_average]
        post_weeks = [receptions, touchdowns]
        ls = pre_weeks + all_weeks + post_weeks

        wb[title].append(ls)


def get_nfl_rush_atts(wb):
    """Retrieve receptions from lineups.com API."""
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/rush/2018/OFF'
    filename = 'nfl_rush_atts.json'

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = pull_data(filename, ENDPOINT)

    # we just want player data
    player_data = data['data']

    # create worksheet
    title = 'RUSH_ATTS'
    header = ['name', 'position', 'team', 'season average', 'week1', 'week2', 'week3', 'week4', 'week5', 'week6',
              'week7', 'week8', 'week9', 'week10', 'week11', 'week12', 'week13', 'week14',
              'week15', 'week16', 'attempts', 'touchdowns']
    create_sheet_header(wb, title, header)

    for d in player_data:
        # TODO rushing_attempt_percentage_by_week
        name = d['name']
        position = d['position']
        team = d['team']
        attempts = d['total']
        weeks = d['weeks']  # dict
        season_average = d['average']
        touchdowns = d['touchdowns']

        # we only care about QB/RB/TE/WR
        if position not in ['QB', 'RB', 'TE', 'WR']:
            continue

        # convert weeks dict to list
        all_weeks = []
        for i in range(0, len(weeks)):
            # if weeks is None, put in blank string
            # 0 would mean they played but didn't get a snap
            if weeks[str(i + 1)] is None:
                all_weeks.append('')
            else:
                all_weeks.append(weeks[str(i + 1)])

        # pad weeks to 16 (a = [])
        # more visual/pythonic
        # a = (a + N * [''])[:N]
        N = 16
        all_weeks = (all_weeks + N * [''])[:N]

        # add three lists together
        pre_weeks = [name, position, team, season_average]
        post_weeks = [attempts, touchdowns]
        ls = pre_weeks + all_weeks + post_weeks

        wb[title].append(ls)


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


def get_dvoa_rankings(wb):
    ENDPOINT = 'https://www.footballoutsiders.com/stats/teamdef'
    filename = 'html_defense.html'

    soup = None
    if not path.isfile(filename):
        print("{} does not exist. Pulling from endpoint [{}]".format(filename, ENDPOINT))
        # send GET request
        r = requests.get(ENDPOINT)
        status = r.status_code

        # if not successful, raise an exception
        if status != 200:
            raise Exception('Requests status != 200. It is: {0}'.format(status))

        # dump html to file to avoid multiple requests
        with open(filename, 'w') as outfile:
            print(r.text, file=outfile)

        soup = BeautifulSoup(r.text, 'html5lib')
    else:
        print("File exists [{}]. Nice!".format(filename))
        # load html from file
        with open(filename, 'r') as html_file:
            soup = BeautifulSoup(html_file, 'html5lib')

    # find all tables in the html
    table = soup.findAll('table')

    if table:
        print("Found {} tables".format(len(table)))

        # create worksheet
        title = 'TEAMDEF'
        wb.create_sheet(title=title)

        defense_stats = table[0]

        # find header
        table_header = defense_stats.find('thead')
        # there is one header row
        header_row = table_header.find('tr')
        # loop through header columns and append to worksheet
        header_cols = header_row.find_all('th')
        header = [ele.text.strip() for ele in header_cols]
        wb[title].append(header)

        # find the rest of the table header_rows
        rows = defense_stats.find_all('tr')
        for row in rows:
            cols = row.find_all('td')
            cols = [ele.text.strip() for ele in cols]
            if cols:
                wb[title].append(cols)

        # separate function for second table
        get_dvoa_recv_rankings(wb, table[1], title)


def get_dvoa_recv_rankings(wb, soup_table, title):
    # VS types of receivers
    def_recv_stats = soup_table
    table_header = def_recv_stats.find('thead')
    header_rows = table_header.find_all('tr')

    # style for merge + center
    al = Alignment(horizontal="center", vertical="center")

    # there are two header rows
    for i, row in enumerate(header_rows):
        header_cols = row.find_all('th')
        header = [ele.text.strip() for ele in header_cols]
        # first header row has some merged cells
        if i == 0:
            # merge + center
            wb[title]['C35'] = header[2]  # vs. WR1
            wb[title].merge_cells('C35:F35')
            style_range(wb[title], 'C35:F35', alignment=al)
            wb[title]['G35'] = header[3]  # vs. WR2
            wb[title].merge_cells('G35:J35')
            style_range(wb[title], 'G35:J35', alignment=al)
            wb[title]['K35'] = header[4]  # vs. OTHER
            wb[title].merge_cells('K35:N35')
            style_range(wb[title], 'K35:N35', alignment=al)
            wb[title]['O35'] = header[5]  # vs. TE
            wb[title].merge_cells('O35:R35')
            style_range(wb[title], 'O35:R35', alignment=al)
            wb[title]['S35'] = header[6]  # vs. RB
            wb[title].merge_cells('S35:V35')
            style_range(wb[title], 'S35:V35', alignment=al)
        elif i == 1:
            wb[title].append(header)
        # for c in cols:
        #     print(c.get_text(strip=True))
        # print(cols)

        # create_sheet_header(wb, title, header)
        # print(header)

    rows = def_recv_stats.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        if cols:
            wb[title].append(cols)


def main():
    fn = 'DKSalaries.csv'
    dest_filename = 'sheet.xlsx'

    # create workbook/worksheet
    wb = Workbook()
    ws1 = wb.active
    ws1.title = 'DKSalaries'

    # guess types (numbers, floats, etc)
    wb.guess_types = True

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

    # pull stats
    get_nfl_receptions(wb)
    get_nfl_targets(wb)
    get_nfl_snaps(wb)
    get_nfl_rush_atts(wb)
    get_dvoa_rankings(wb)

    # save workbook (.xlsx file)
    wb.save(filename=dest_filename)


if __name__ == "__main__":
    main()
