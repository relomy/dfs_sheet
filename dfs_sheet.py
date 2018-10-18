"""Create DFS spreadsheet from stats """

import json
import re
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font, colors
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
# from openpyxl.cell import get_column_letter
from os import path, makedirs


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
    fn = 'nfl_snaps.json'
    dir = 'sources'
    filename = path.join(dir, fn)

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = pull_data(filename, ENDPOINT)

    if data is None:
        raise Exception('Failed to pull data from API or file.')

    player_data = data['data']

    # create worksheet
    title = 'SNAPS'
    header = ['name', 'position', 'team', 'season average', 'week1', 'week2', 'week3', 'week4', 'week5', 'week6',
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

        # remove '.' from name
        name = name.replace('.', '')

        # convert weeks dict to list
        all_weeks = conv_weeks_to_padded_list(weeks)

        # add three lists together
        pre_weeks = [name, position, team, season_average]
        # post_weeks = [targets, average, recv_touchdowns]
        ls = pre_weeks + all_weeks

        wb[title].append(ls)


def get_nfl_targets(wb):
    """Retrieve targets from lineups.com API."""
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/targets/2018/OFF'
    fn = 'nfl_targets.json'
    dir = 'sources'
    filename = path.join(dir, fn)

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

        # remove '.' from name
        name = name.replace('.', '')

        # convert weeks dict to list
        all_weeks = conv_weeks_to_padded_list(weeks)

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
    fn = 'nfl_receptions.json'
    dir = 'sources'
    filename = path.join(dir, fn)

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = pull_data(filename, ENDPOINT)

    # we just want player data
    player_data = data['data']

    # create worksheet
    title = 'RECEPTIONS'
    header = ['name', 'position', 'team', 'season average', 'week1', 'week2', 'week3', 'week4', 'week5', 'week6',
              'week7', 'week8', 'week9', 'week10', 'week11', 'week12', 'week13', 'week14',
              'week15', 'week16', 'receptions', 'touchdowns']
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

        # remove '.' from name
        name = name.replace('.', '')

        # convert weeks dict to list
        all_weeks = conv_weeks_to_padded_list(weeks)

        # add three lists together
        pre_weeks = [name, position, team, season_average]
        post_weeks = [receptions, touchdowns]
        ls = pre_weeks + all_weeks + post_weeks

        wb[title].append(ls)


def get_nfl_rush_atts(wb):
    """Retrieve receptions from lineups.com API."""
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/rush/2018/OFF'
    fn = 'nfl_rush_atts.json'
    dir = 'sources'
    filename = path.join(dir, fn)

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

        # we only care about QB/RB/WR
        if position not in ['QB', 'RB', 'WR']:
            continue

        # remove '.' from name
        name = name.replace('.', '')

        # convert weeks dict to list
        all_weeks = conv_weeks_to_padded_list(weeks)

        # add three lists together
        pre_weeks = [name, position, team, season_average]
        post_weeks = [attempts, touchdowns]
        ls = pre_weeks + all_weeks + post_weeks

        wb[title].append(ls)


def conv_weeks_to_padded_list(weeks):
    """Convert weeks dict or list to padded list (16 weeks)."""
    all_weeks = []
    if isinstance(weeks, list):
        for week in weeks:
            if week is None:
                all_weeks.append('')
            else:
                all_weeks.append(week)
    elif isinstance(weeks, dict):
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

    return all_weeks


def get_vegas_rg(wb):
    ENDPOINT = 'https://rotogrinders.com/schedules/nfl'

    fn = 'vegas_script.html'
    dir = 'sources'
    filename = path.join(dir, fn)

    # create worksheet
    title = 'VEGAS'
    header = ['Time', 'Team', 'Opponent', 'Line', 'MoneyLine', 'Over/Under', 'Projected Points', 'Projected Points Change']
    create_sheet_header(wb, title, header)

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

    # find script(s) in the html
    script = soup.findAll('script')

    js_vegas_data = script[11].string

    # replace dumb names
    js_vegas_data = js_vegas_data.replace('KCC', 'KC')
    js_vegas_data = js_vegas_data.replace('JAC', 'JAX')

    pattern = re.compile(r'data = (.*);')

    json_str = pattern.search(js_vegas_data).group(1)
    vegas_json = json.loads(json_str)
    for matchup in vegas_json:
        wb[title].append([
            matchup['time']['display'],
            matchup['team'],
            matchup['opponent'],
            matchup['line'],
            matchup['moneyline'],
            matchup['overunder'],
            matchup['projected'],
            matchup['projectedchange']['value']
        ])


def get_dvoa_rankings(wb):
    ENDPOINT = 'https://www.footballoutsiders.com/stats/teamdef'
    fn = 'html_defense.html'
    dir = 'sources'
    filename = path.join(dir, fn)

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

    # find all tables (3) in the html
    table = soup.findAll('table')

    if table:
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


def get_oline_rankings(wb):
    ENDPOINT = 'https://www.footballoutsiders.com/stats/ol'
    fn = 'html_oline.html'
    dir = 'sources'
    filename = path.join(dir, fn)

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

    # find all tables (2) in the html
    table = soup.findAll('table')

    if table:
        # create worksheet
        title = 'OLINE'
        wb.create_sheet(title=title)

        oline_stats = table[0]

        # find header
        table_header = oline_stats.find('thead')
        # there is one header row
        header_row = table_header.find('tr')
        # loop through header columns and append to worksheet
        header_cols = header_row.find_all('th')
        header = [ele.text.strip() for ele in header_cols]
        wb[title].append(header)

        # find the rest of the table header_rows
        rows = oline_stats.find_all('tr')
        for row in rows:
            cols = row.find_all('td')
            cols = [ele.text.strip() for ele in cols]
            if cols:
                wb[title].append(cols)


def get_dline_rankings(wb):
    ENDPOINT = 'https://www.footballoutsiders.com/stats/dl'
    fn = 'html_dline.html'
    dir = 'sources'
    filename = path.join(dir, fn)

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

    # find all tables (2) in the html
    table = soup.findAll('table')

    if table:
        # create worksheet
        title = 'DLINE'
        wb.create_sheet(title=title)

        oline_stats = table[0]

        # find header
        table_header = oline_stats.find('thead')
        # there is one header row
        header_row = table_header.find('tr')
        # loop through header columns and append to worksheet
        header_cols = header_row.find_all('th')
        header = [ele.text.strip() for ele in header_cols]
        wb[title].append(header)

        # find the rest of the table header_rows
        rows = oline_stats.find_all('tr')
        for row in rows:
            cols = row.find_all('td')
            cols = [ele.text.strip() for ele in cols]
            if cols:
                wb[title].append(cols)


def get_qb_stats(wb):
    ENDPOINT = 'https://www.footballoutsiders.com/stats/qb'
    fn = 'html_qb.html'
    dir = 'sources'
    filename = path.join(dir, fn)

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

    # find all tables (3) in the html
    table = soup.findAll('table')

    if table:
        # create worksheet
        title = 'QB_STATS'
        wb.create_sheet(title=title)

        for t in table:
            qb_stats = t

            # find header
            table_header = qb_stats.find('thead')
            # there is one header row
            header_row = table_header.find('tr')
            # loop through header columns and append to worksheet
            header_cols = header_row.find_all('th')
            header = [ele.text.strip() for ele in header_cols]
            wb[title].append(header)

            # find the rest of the table header_rows
            rows = qb_stats.find_all('tr')
            for row in rows:
                cols = row.find_all('td')
                cols = [ele.text.strip() for ele in cols]
                if cols:
                    wb[title].append(cols)


def fpros_ecr(wb, position):
    if position == 'QB' or position == 'DST':
        ENDPOINT = 'https://www.fantasypros.com/nfl/rankings/{}.php'.format(position.lower())
    else:
        ENDPOINT = 'https://www.fantasypros.com/nfl/rankings/ppr-{}.php'.format(position.lower())

    fn = 'ecr_{}.html'.format(position)
    dir = 'sources'
    filename = path.join(dir, fn)

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

    # find all tables (2) in the html
    table = soup.find('table', id='rank-data')

    if table:
        # create worksheet
        title = '{0}_ECR'.format(position)
        wb.create_sheet(title=title)

        # # find header
        table_header = table.find('thead')
        # there is one header row
        header_row = table_header.find('tr')
        # loop through header columns and append to worksheet
        header_cols = header_row.find_all('th')
        header = [ele.text.strip() for ele in header_cols]
        wb[title].append(header)

        # find the rest of the table header_rows
        rows = table.find_all('tr')
        for row in rows:
            cols = row.find_all('td')
            # cols = [ele.text.strip() for ele in cols]
            # change from list comp for just fpros
            new_cols = []
            for ele in cols:
                txt = ele.text.strip()
                # replace JAX
                txt = txt.replace('JAC', 'JAX')
                # remove periods (T.J. Yeldon, T.Y. Hilton)
                txt = txt.replace('.', '')
                # really? just to fix mitchell?
                if position == 'QB':
                    txt = txt.replace('Mitch', 'Mitchell')
                new_cols.append(txt)
            if cols:
                wb[title].append(new_cols)


def position_tab(wb, values, title):
    # create positional tab if it does not exist
    # and set header(s)
    if title not in wb.sheetnames:
        wb.create_sheet(title=title)

        # style for merge + center
        al = Alignment(horizontal="center", vertical="center")

        # second header
        all_positions_header = [
            'Position', 'Name', 'Opp', 'Salary', 'Salary%', 'Abbv',
            'Implied Total', 'O/U', 'Line'
        ]

        # set row height
        wb[title].row_dimensions[2].height = 40

        # more header fields based on position
        position_fields = []
        if title == 'QB':
            top_lvl_header(wb, title, 'DK', 'D', 1, 'FF000000')
            top_lvl_header(wb, title, 'VEGAS', 'G', 2, 'FFFFC000')
            top_lvl_header(wb, title, 'MATCHUP', 'J', 2, 'FFED7D31')
            top_lvl_header(wb, title, 'PRESSURE', 'M', 1, 'FF5B9BD5')
            top_lvl_header(wb, title, 'RANKINGS', 'O', 1, 'FF70AD47')

            position_fields = [
                'Rushing Yards', 'DYAR', 'QBR', 'O-Line Sack%', 'D-Line Sack%',
                'Average PPG', 'ECR', 'ECR Data'
            ]
        elif title == 'RB':
            # Starting with D1
            # DK, DK%, blank, VEGASx3, MATCHUPx4, SEASON,x3, LAST WEEKx3, RANKINGSx2
            # top header
            # bold, white font
            font = Font(b=True, color="FFFFFFFF")

            # top level header
            top_lvl_header(wb, title, 'DK', 'D', 1, 'FF000000')
            top_lvl_header(wb, title, 'VEGAS', 'G', 2, 'FFFFC000')
            top_lvl_header(wb, title, 'MATCHUP', 'J', 3, 'FFED7D31')
            top_lvl_header(wb, title, 'SEASON', 'N', 2, 'FF5B9BD5')
            top_lvl_header(wb, title, 'LAST WEEK', 'Q', 2, 'FF4472C4')
            top_lvl_header(wb, title, 'RANKINGS', 'T', 1, 'FF70AD47')

            position_fields = [
                'Run DVOA', 'Pass DVOA', 'O-Line', 'D-Line', 'Snap%', 'Rush ATTs',
                'Targets', 'Snap%', 'Rush ATTs', 'Targets', 'Average PPG', 'ECR', 'ECR Data'
            ]
        elif title == 'WR':
            # Starting with D1
            # DK, DK%, blank, VEGASx3, MATCHUPx4, SEASON,x3, LAST WEEKx3, RANKINGSx2
            # top header
            top_lvl_header(wb, title, 'DK', 'D', 1, 'FF000000')
            top_lvl_header(wb, title, 'VEGAS', 'G', 2, 'FFFFC000')
            top_lvl_header(wb, title, 'MATCHUP', 'J', 2, 'FFED7D31')
            top_lvl_header(wb, title, 'SEASON', 'M', 1, 'FF5B9BD5')
            top_lvl_header(wb, title, 'LAST WEEK', 'O', 1, 'FF4472C4')
            top_lvl_header(wb, title, 'RANKINGS', 'Q', 1, 'FF70AD47')

            position_fields = [
                'DVOA', 'vs. WR1', 'vs. WR2', 'Snap%', 'Targets', 'Snap%', 'Targets',
                'Average PPG', 'ECR', 'ECR Data'
            ]
        elif title == 'TE':
            top_lvl_header(wb, title, 'DK', 'D', 1, 'FF000000')
            top_lvl_header(wb, title, 'VEGAS', 'G', 2, 'FFFFC000')
            top_lvl_header(wb, title, 'MATCHUP', 'J', 1, 'FFED7D31')
            top_lvl_header(wb, title, 'SEASON', 'L', 1, 'FF5B9BD5')
            top_lvl_header(wb, title, 'LAST WEEK', 'N', 1, 'FF4472C4')
            top_lvl_header(wb, title, 'RANKINGS', 'P', 1, 'FF70AD47')

            position_fields = [
                'DVOA', 'vs. TE', 'Snap%', 'Targets', 'Snap%', 'Targets', 'Average PPG',
                'ECR', 'ECR Data'
            ]
        elif title == 'DST':
            top_lvl_header(wb, title, 'DK', 'D', 1, 'FF000000')
            top_lvl_header(wb, title, 'VEGAS', 'G', 2, 'FFFFC000')
            top_lvl_header(wb, title, 'RANKINGS', 'O', 1, 'FF70AD47')

            position_fields = [
                'Average PPG', 'ECR'
            ]

        # find max row to append
        append_row = wb[title].max_row + 1
        header = all_positions_header + position_fields

        # change row font and alignment
        font = Font(b=True, color="FF000000")
        al = Alignment(horizontal="center", vertical="center", wrapText=True)

        # just set for row range
        rng = "{0}:{1}".format(2, 2)
        for cell in wb[title][rng]:
            cell.font = font
            cell.alignment = al

        for i, field in enumerate(header):
            wb[title].cell(row=append_row, column=i + 1, value=field)

        # freeze header
        wb[title].freeze_panes = "O3"

    keys = ['pos', 'name_id', 'name', 'id', 'roster_pos', 'salary', 'matchup', 'abbv', 'avg_ppg']
    stats_dict = dict(zip(keys, values))
    stats_dict['salary_perc'] = "{0:.1%}".format(float(stats_dict['salary']) / 50000)

    # 'fix' name to remove extra stuff like Jr or III (Todd Gurley II for example)
    name = ' '.join(stats_dict['name'].split(' ')[:2])
    # also remove periods (T.J. Yeldon for example)
    name = name.replace('.', '')
    stats_dict['name'] = name

    # find opp, opp_excel, and game_time
    home_at_away, game_time = stats_dict['matchup'].split(' ', 1)
    stats_dict['game_time'] = game_time
    home_team, away_team = home_at_away.split('@')
    if stats_dict['abbv'] == home_team:
        stats_dict['opp'] = away_team
        stats_dict['opp_excel'] = "vs. {}".format(away_team)
    else:
        stats_dict['opp'] = home_team
        stats_dict['opp_excel'] = "at {}".format(home_team)

    # find max row to append
    append_row = wb[title].max_row + 1

    # insert rows of data
    all_positions_fields = [
        stats_dict['pos'],
        stats_dict['name'],
        stats_dict['opp_excel'],
        stats_dict['salary'],
        stats_dict['salary_perc'],
        stats_dict['abbv'],
        bld_excel_formula('VEGAS', '$G$2:$G$29', '$F', append_row, '$B$2:$B$29'),  # implied total
        bld_excel_formula('VEGAS', '$F$2:$F$29', '$F', append_row, '$B$2:$B$29'),  # over/under
        bld_excel_formula('VEGAS', '$D$2:$D$29', '$F', append_row, '$B$2:$B$29'),  # line
    ]

    # more header fields based on position
    positional_fields = []

    # get max_row from position ECR tab
    max_row = wb[title + '_ECR'].max_row
    if title == 'QB':
        positional_fields = [
            # rushing yards
            bld_excel_formula('QB_STATS', '$K$44:$K$82', '$B', append_row, '$A$44:$A$82', qb_stats=True),
            # DYAR
            bld_excel_formula('QB_STATS', '$C$2:$C$42', '$B', append_row, '$A$2:$A$42', qb_stats=True),
            # QBR
            bld_excel_formula('QB_STATS', '$J$2:$J$35', '$B', append_row, '$A$2:$A$35', qb_stats=True),
            # o-line
            bld_excel_formula('OLINE', 'P$2:$P$35', '$F', append_row, '$B$2:$B$33'),
            # d-line
            bld_excel_formula('DLINE', 'P$2:$P$35', '$C', append_row, '$B$2:$B$33', right=True),
            # average PPG
            stats_dict['avg_ppg'],
            # =OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,1) # cell to the right
            # ECR
            '=RANK(Q{0}, $Q$3:$Q$53,1)'.format(append_row),
            # ECR DATA
            bld_excel_formula('QB_ECR', '$A$2:$A${}'.format(max_row), '$B', append_row, '$C$2:$C${}'.format(max_row)),
        ]
        # style column L & M (pressure %) with %/decimals
        for cell in wb[title]['M']:
            cell.number_format = '##0.0%'
        for cell in wb[title]['N']:
            cell.number_format = '##0.0%'

        # hide column Q (ECR Data)
        wb[title].column_dimensions['Q'].hidden = True
    elif title == 'RB':
        max_row = wb[title + '_ECR'].max_row
        positional_fields = [
            # run dvoa
            bld_excel_formula('TEAMDEF', 'J$2:$J$33', '$C', append_row, '$B$2:$B$33', right=True),
            # pass dvoa (vs. RB)
            bld_excel_formula('TEAMDEF', 'T$37:$T$68', '$C', append_row, '$B$37:$B$68', right=True),
            # o line
            bld_excel_formula('OLINE', '$C$2:$C$33', '$F', append_row, '$B$2:$B$33'),
            # d line
            bld_excel_formula('DLINE', '$C$2:$C$33', '$C', append_row, '$B$2:$B$33', right=True),
            # season snap%
            bld_excel_formula('SNAPS', '$D$2:$D$449', '$B', append_row, '$A$2:$A$449'),
            # season rush atts
            bld_excel_formula('RUSH_ATTS', '$D$2:$D$449', '$B', append_row, '$A$2:$A$449'),
            # season targets
            bld_excel_formula('TARGETS', '$D$2:$D$449', '$B', append_row, '$A$2:$A$449'),
            # week snap% (week6)
            bld_excel_formula('SNAPS', '$J$2:$J$449', '$B', append_row, '$A$2:$A$449', week=True),
            # week rush atts (week 6)
            bld_excel_formula('RUSH_ATTS', '$J$2:$J$449', '$B', append_row, '$A$2:$A$449', week=True),
            # week targets (week 6)
            bld_excel_formula('TARGETS', '$J$2:$J$449', '$B', append_row, '$A$2:$A$449', week=True),
            # average PPG
            stats_dict['avg_ppg'],
            # ECR
            '=RANK(V{0}, $V$3:$V$69,1)'.format(append_row),
            # ECR Data
            bld_excel_formula('RB_ECR', '$A$2:$A${}'.format(max_row), '$B', append_row, '$C$2:$C${}'.format(max_row)),
        ]
        # hide column V (ECR Data)
        wb[title].column_dimensions['V'].hidden = True
    elif title == 'WR':
        positional_fields = [
            # pass dvoa
            bld_excel_formula('TEAMDEF', '$H$2:$H$34', '$C', append_row, '$B$2:$B$34', right=True),
            # vs. WR1
            bld_excel_formula('TEAMDEF', '$D$37:$D$68', '$C', append_row, '$B$37:$B$68', right=True),
            # vs. WR2
            bld_excel_formula('TEAMDEF', '$H$37:$H$68', '$C', append_row, '$B$37:$B$68', right=True),
            # season snap%
            bld_excel_formula('SNAPS', '$D$2:$D$449', '$B', append_row, '$A$2:$A$449'),
            # season targets
            bld_excel_formula('TARGETS', '$D$2:$D$449', '$B', append_row, '$A$2:$A$449'),
            # week snap% (week6)
            bld_excel_formula('SNAPS', '$J$2:$J$449', '$B', append_row, '$A$2:$A$449', week=True),
            # week targets (week 6)
            bld_excel_formula('TARGETS', '$J$2:$J$449', '$B', append_row, '$A$2:$A$449', week=True),
            # average PPG
            stats_dict['avg_ppg'],
            # ECR
            '=RANK(S{0}, $S$3:$S$94,1)'.format(append_row),
            # ECR Data
            bld_excel_formula('WR_ECR', '$A$2:$A${}'.format(max_row), '$B', append_row, '$C$2:$C${}'.format(max_row)),
        ]
        # hide column S (ECR Data)
        wb[title].column_dimensions['S'].hidden = True
    elif title == 'TE':
        positional_fields = [
            # pass dvoa
            bld_excel_formula('TEAMDEF', '$H$2:$H$34', '$C', append_row, '$B$2:$B$34', right=True),
            # vs. TE
            bld_excel_formula('TEAMDEF', '$P$37:$P$68', '$C', append_row, '$B$37:$B$68', right=True),
            # season snap%
            bld_excel_formula('SNAPS', '$D$2:$D$449', '$B', append_row, '$A$2:$A$449'),
            # season targets
            bld_excel_formula('TARGETS', '$D$2:$D$449', '$B', append_row, '$A$2:$A$449'),
            # week snap% (week6)
            bld_excel_formula('SNAPS', '$J$2:$J$449', '$B', append_row, '$A$2:$A$449', week=True),
            # week targets (week 6)
            bld_excel_formula('TARGETS', '$J$2:$J$449', '$B', append_row, '$A$2:$A$449', week=True),
            # average PPG
            stats_dict['avg_ppg'],
            # ECR
            '=RANK(R{0}, $R$3:$R$52,1)'.format(append_row),
            # ECR Data
            bld_excel_formula('TE_ECR', '$A$2:$A${}'.format(max_row), '$B', append_row, '$C$2:$C${}'.format(max_row)),
        ]
        # hide column r (ECR Data)
        wb[title].column_dimensions['R'].hidden = True
    elif title == 'DST':
        positional_fields = [
            # average PPG
            stats_dict['avg_ppg'],
            # ECR
            '=RANK(L{0}, $L$3:$L$52,1)'.format(append_row),
            # ECR Data
            bld_excel_formula('DST_ECR', '$A$2:$A${}'.format(max_row), '$F', append_row, '$C$2:$C{}'.format(max_row), dst=True),
        ]

        # hide column l (ECR Data)
        wb[title].column_dimensions['L'].hidden = True
    row = all_positions_fields + positional_fields

    # center all cells horzitontally/vertically in row
    for i, text in enumerate(row, start=1):
        nice = wb[title].cell(row=append_row, column=i, value=text)
        al = Alignment(horizontal="center", vertical="center")
        nice.alignment = al

    # style column D (salary) with currency
    for cell in wb[title]['D']:
        cell.number_format = '$#,##0_);($#,##0)'

    # style column E (salary %) with %/decimals
    for cell in wb[title]['E']:
        cell.number_format = '##0.0%'

    # hide column F (abbv)
    wb[title].column_dimensions['F'].hidden = True


def top_lvl_header(wb, title, text, start_col, length, color):
    # style for merge + center
    al = Alignment(horizontal="center", vertical="center")
    # bold font
    font = Font(b=True, color="FFFFFFFF")

    # set cell to start merge + insert text
    cell = "{0}1".format(start_col)
    wb[title][cell] = text
    # set range to format merged cells
    fmt_range = "{0}1:{1}1".format(start_col, chr(ord(start_col) + length))
    style_range(wb[title], fmt_range, font=font, fill=PatternFill(patternType="solid", fgColor=color), alignment=al)


def bld_excel_formula(title, rtrn_range, match, row, match_range, week=False, right=False, qb_stats=False, dst=False):
    # '=INDEX(OLINE!$C$2:$C$33,MATCH($F{0},OLINE!$B$2:$B$33,0))'.format(append_row),

    # use RIGHT for splitting the opponent. IE JAX for vs. JAX
    if right:
        base_formula = 'INDEX({0}!{1}, MATCH(RIGHT({2}{3}, LEN({2}{3}) - SEARCH(" ",{2}{3},1)) & "*", {0}!{4},0))'.format(
            title, rtrn_range, match, row, match_range)
    elif qb_stats:
        base_formula = 'INDEX({0}!{1}, MATCH(LEFT({2}{3}, 1) & "*" & RIGHT({2}{3}, LEN({2}{3}) - SEARCH(" ",{2}{3},1)) & "*", {0}!{4},0))'.format(
            title, rtrn_range, match, row, match_range)
    elif dst:
        base_formula = 'INDEX({0}!{1}, MATCH("*(" & {2}{3} & "*", {0}!{4},0))'.format(
            title, rtrn_range, match, row, match_range)
    else:
        base_formula = 'INDEX({0}!{1}, MATCH({2}{3} & "*", {0}!{4},0))'.format(
            title, rtrn_range, match, row, match_range)

    # if we are looking at weekly stats, blank should be blank, not zero
    if week:
        formula = 'IF(ISBLANK({0}), " ", {0})'.format(base_formula)
    else:
        formula = base_formula

    return "=" + formula


def style_ranges(wb):
    # define colors for colorscale (from excel)
    red = 'F8696B'
    yellow = 'FFEB84'
    green = '63BE7B'

    for title in ['QB', 'RB', 'WR', 'TE', 'DST']:
        ws = wb[title]
        # add filter/sort. excel will not automatically do it!
        # filter_range = "{0}:{1}".format('D2', ws.max_row)
        # ws.auto_filter.ref = filter_range
        # sort_range = "{0}:{1}".format('D3', ws.max_row)

        # ws.auto_filter.add_sort_condition(sort_range)
        # bigger/positive = green, smaller/negative = red
        green_to_red_headers = [
            'Implied Total', 'O/U', 'Run DVOA', 'Pass DVOA', 'DVOA', 'vs. WR1', 'vs. WR2',
            'O-Line', 'Snap%', 'Rush ATTs', 'Targets', 'vs. TE', 'D-Line Sack%', 'Average PPG',
            'Rushing Yards', 'DYAR', 'QBR'
        ]
        green_to_red_rule = ColorScaleRule(start_type='min', start_color=red,
                                           mid_type='percentile', mid_value=50, mid_color=yellow,
                                           end_type='max', end_color=green)
        # bigger/positive = red, smaller/negative = green
        red_to_green_headers = [
            'Line', 'D-Line', 'O-Line Sack%', 'ECR'
        ]
        red_to_green_rule = ColorScaleRule(start_type='min', start_color=green,
                                           mid_type='percentile', mid_value=50, mid_color=yellow,
                                           end_type='max', end_color=red)
        # color ranges
        for i in range(1, ws.max_column + 1):
            if ws.cell(row=2, column=i).value in green_to_red_headers:
                column_letter = get_column_letter(i)
                # color range (green to red)
                cell_rng = "{0}{1}:{2}".format(column_letter, '3', ws.max_row)
                # print("[{}] Coloring {} [{} - {}] green_to_red".format(title, ws.cell(row=2, column=i).value, ws.cell(row=2, column=i), cell_rng))
                wb[title].conditional_formatting.add(cell_rng, green_to_red_rule)
            elif ws.cell(row=2, column=i).value in red_to_green_headers:
                column_letter = get_column_letter(i)
                # color range (red to green)
                cell_rng = "{0}{1}:{2}".format(column_letter, '3', ws.max_row)
                # print("[{}] Coloring {} [{} - {}] red_to_green".format(title, ws.cell(row=2, column=i).value, ws.cell(row=2, column=i), cell_rng))
                wb[title].conditional_formatting.add(cell_rng, red_to_green_rule)

        # set column widths
        column_widths = [8, 20, 10, 8, 8, 8, 8, 8, 8, 8, 8, 8]
        for i, column_width in enumerate(column_widths):
            ws.column_dimensions[get_column_letter(i + 1)].width = column_width


def remove_rows_without_ecr(wb):
    for title in ['QB']:
        ws = wb[title]

        name_col = 'B'
        for cell in ws[name_col]:
            # skip first row (None)
            if cell.row <= 2:
                continue

            # guy we are looking for
            name = cell.value

            # get ECR sheet
            ecr_ws = wb[title + '_ECR']
            search_col = 'C'

            # search ECR sheet for guy
            bool = bool_found_player_in_ecr_tab(ecr_ws[search_col], name)

            if bool is False:
                print("Did not find {}.. will delete row {}".format(name, cell.row))
                ws.delete_rows(cell.row)


def bool_found_player_in_ecr_tab(ws_column, name):
    # loop through cells in column
    for c in ws_column:
        # if cell is empty, continue
        if c.value is None:
            continue

        # if name is found, move along
        if name in c.value:
            return True
    return False


def order_sheets(wb):
    # QB, RB, WR, TE, DST first
    order = [wb.worksheets.index(wb[i]) for i in ['QB', 'RB', 'WR', 'TE', 'DST']]
    # order = [wb.worksheets.index(wb['QB']),
    #          wb.worksheets.index(wb['RB']),
    #          wb.worksheets.index(wb['WR']),
    #          wb.worksheets.index(wb['TE']),
    #          wb.worksheets.index(wb['DST'])]

    order.extend(list(set(range(len(wb._sheets))) - set(order)))
    wb._sheets = [wb._sheets[i] for i in order]


def check_name_in_ecr(wb, position, name):
    # get ECR sheet
    ecr_ws = wb[position + '_ECR']
    search_col = 'C'

    # search ECR sheet for guy
    return bool_found_player_in_ecr_tab(ecr_ws[search_col], name)


def main():
    fn = 'DKSalaries_week7_full.csv'
    dest_filename = 'sheet.xlsx'

    # create workbook/worksheet
    wb = Workbook()
    ws1 = wb.active
    ws1.title = 'DEL'

    # guess types (numbers, floats, etc)
    wb.guess_types = True

    # make sources dir if it does not exist
    directory = 'sources'
    if not path.exists(directory):
        makedirs(directory)

    # pull stats from fantasypros.com
    fpros_ecr(wb, 'QB')
    fpros_ecr(wb, 'RB')
    fpros_ecr(wb, 'WR')
    fpros_ecr(wb, 'TE')
    fpros_ecr(wb, 'DST')

    with open(fn, 'r') as f:
        # read entire file into memory
        lines = f.readlines()

        for i, line in enumerate(lines):
            # skip header
            if i == 0:
                continue

            fields = line.rstrip().split(',')

            # check if player has ECR
            position = fields[0]
            name = fields[2]

            # 'fix' name to remove extra stuff like Jr or III (Todd Gurley II for example)
            name = ' '.join(name.split(' ')[:2])
            # also remove periods (T.J. Yeldon for example)
            name = name.replace('.', '')

            if position == 'DST':
                name = fields[7]

            # if player does not exist, skip
            if check_name_in_ecr(wb, position, name) is False:
                # print("Could not find {} [{}]".format(name, position))
                continue

            position_tab(wb, fields, fields[0])

    # pull stats from lineups.com
    get_nfl_receptions(wb)
    get_nfl_targets(wb)
    get_nfl_snaps(wb)
    get_nfl_rush_atts(wb)
    # pull stats from footballoutsiders.com
    get_dvoa_rankings(wb)
    get_oline_rankings(wb)
    get_dline_rankings(wb)
    get_qb_stats(wb)
    # pull vegas stats from rotogrinders.com
    get_vegas_rg(wb)

    # color ranges
    style_ranges(wb)
    # remove rows without ECR ranking (either out or useless)
    # remove_rows_without_ecr(wb)

    # order sheets
    # wb._sheets =[wb._sheets[i] for i in myorder]
    order_sheets(wb)

    # save workbook (.xlsx file)
    wb.remove(ws1)  # remove blank worksheet
    wb.save(filename=dest_filename)

    # remove rows without an ECR ranking (likely out or useless))
    # wb_data_only = load_workbook(dest_filename, data_only=True)


if __name__ == "__main__":
    main()
