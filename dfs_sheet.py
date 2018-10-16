import json
import dirtyjson
import re
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font, colors
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
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
        all_weeks = conv_weeks_to_padded_list(weeks)

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
        all_weeks = conv_weeks_to_padded_list(weeks)

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


def get_vegas_rg(wb):
    ENDPOINT = 'https://rotogrinders.com/schedules/nfl'
    filename = 'vegas_script.html'

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
    filename = 'html_oline.html'

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
    filename = 'html_dline.html'

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
        if title == 'RB':


            # Starting with D1
            # DK, DK%, blank, VEGASx3, MATCHUPx4, SEASON,x3, LAST WEEKx3, RANKINGSx2
            # top header
            # bold, white font
            font = Font(b=True, color="FFFFFFFF")

            # top level header
            wb[title]['D1'] = 'DK'
            style_range(wb[title], 'D1:F1', font=font, fill=PatternFill(patternType="solid", fgColor="FF000000"), alignment=al)
            wb[title]['G1'] = 'VEGAS'
            style_range(wb[title], 'G1:I1', font=font, fill=PatternFill(patternType="solid", fgColor="FFFFC000"), alignment=al)
            wb[title]['J1'] = 'MATCHUP'
            style_range(wb[title], 'J1:M1', font=font, fill=PatternFill(patternType="solid", fgColor="FFED7D31"), alignment=al)
            wb[title]['N1'] = 'SEASON'
            style_range(wb[title], 'N1:P1', font=font, fill=PatternFill(patternType="solid", fgColor="FF5B9BD5"), alignment=al)
            wb[title]['Q1'] = 'LAST WEEK'
            style_range(wb[title], 'Q1:S1', font=font, fill=PatternFill(patternType="solid", fgColor="FF4472C4"), alignment=al)
            wb[title]['T1'] = 'RANKINGS'
            style_range(wb[title], 'T1:U1', font=font, fill=PatternFill(patternType="solid", fgColor="FF70AD47"), alignment=al)

            position_fields = [
                'Run DVOA', 'Pass DVOA', 'O-Line', 'D-Line', 'Snap%', 'Rush ATTs',
                'Targets', 'Snap%', 'Rush ATTs', 'Targets', 'Average PPG', 'ECR'
            ]
        elif title == 'WR':
            # set row height
            wb[title].row_dimensions[2].height = 40
            # Starting with D1
            # DK, DK%, blank, VEGASx3, MATCHUPx4, SEASON,x3, LAST WEEKx3, RANKINGSx2
            # top header
            wb[title]['D1'] = 'DK'
            style_range(wb[title], 'D1:F1', alignment=al)
            wb[title]['G1'] = 'VEGAS'
            style_range(wb[title], 'G1:I1', alignment=al)
            wb[title]['J1'] = 'MATCHUP'
            style_range(wb[title], 'J1:L1', alignment=al)
            wb[title]['M1'] = 'SEASON'
            style_range(wb[title], 'M1:N1', alignment=al)
            wb[title]['O1'] = 'LAST WEEK'
            style_range(wb[title], 'O1:P1', alignment=al)
            wb[title]['Q1'] = 'RANKINGS'
            style_range(wb[title], 'Q1:R1', alignment=al)

            position_fields = [
                'DVOA', 'WR1', 'WR2', 'Snap%', 'Targets', 'Snap%', 'Targets', 'Average PPG', 'ECR'
            ]

        header = all_positions_header + position_fields

        append_row = wb[title].max_row + 1

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

    keys = ['pos', 'name_id', 'name', 'id', 'roster_pos', 'salary', 'matchup', 'abbv', 'avg_ppg']
    stats_dict = dict(zip(keys, values))
    stats_dict['salary_perc'] = "{0:.1%}".format(float(stats_dict['salary']) / 50000)

    # 'fix' name to remove extra stuff like Jr or III
    name = ' '.join(stats_dict['name'].split(' ')[:2])
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

    append_row = wb[title].max_row + 1

    # vegas formula OU
    # =INDEX(Vegas!$G$2:$G$29,MATCH($E3 & "*",Vegas!$B$2:$B$29,0))
    # insert rows of data
    all_positions_fields = [
        stats_dict['pos'],
        stats_dict['name'],
        stats_dict['opp_excel'],
        stats_dict['salary'],
        stats_dict['salary_perc'],
        stats_dict['abbv'],
        '=INDEX(Vegas!$G$2:$G$29,MATCH($F{} & "*",Vegas!$B$2:$B$29,0))'.format(append_row),  # implied total
        '=INDEX(Vegas!$F$2:$F$29,MATCH($F{} & "*",Vegas!$B$2:$B$29,0))'.format(append_row),  # over/under
        '=INDEX(Vegas!$D$2:$D$29,MATCH($F{} & "*",Vegas!$B$2:$B$29,0))'.format(append_row)   # line
    ]

    # more header fields based on position
    positional_fields = []
    if title == 'RB':
        positional_fields = [
            # run dvoa
            '=INDEX(TEAMDEF!$J$2:$J$34,MATCH(RIGHT($C{0}, LEN($C{0}) - SEARCH(" ",$C{0},1)),TEAMDEF!$B$2:$B$34,0))'.format(append_row),
            # pass dvoa (vs. RB)
            '=INDEX(TEAMDEF!$T$37:$T$68,MATCH(RIGHT($C{0}, LEN($C{0}) - SEARCH(" ",$C{0},1)),TEAMDEF!$B$37:$B$68,0))'.format(append_row),
            # o line
            '=INDEX(OLINE!$C$2:$C$33,MATCH($F{0},OLINE!$B$2:$B$33,0))'.format(append_row),
            # d line
            '=INDEX(DLINE!$C$2:$C$33,MATCH(RIGHT($C{0}, LEN($C{0}) - SEARCH(" ",$C{0},1)),DLINE!$B$2:$B$33,0))'.format(append_row),
            # season snap%
            '=INDEX(SNAPS!$D$2:$D$448,MATCH($B{0} & "*",SNAPS!$A$2:$A$448,0))'.format(append_row),
            # season rush atts
            # season targets
        ]
    elif title == 'WR':
        positional_fields = [
            # pass dvoa
            '=INDEX(TEAMDEF!$H$2:$H$34,MATCH(RIGHT($C{0}, LEN($C{0}) - SEARCH(" ",$C{0},1)),TEAMDEF!$B$2:$B$34,0))'.format(append_row),
            # vs. WR1
            '=INDEX(TEAMDEF!$D$37:$D$68,MATCH(RIGHT($C{0}, LEN($C{0}) - SEARCH(" ",$C{0},1)),TEAMDEF!$B$37:$B$68,0))'.format(append_row),
            # vs. WR2
            '=INDEX(TEAMDEF!$H$37:$H$68,MATCH(RIGHT($C{0}, LEN($C{0}) - SEARCH(" ",$C{0},1)),TEAMDEF!$B$37:$B$68,0))'.format(append_row),
        ]

    row = all_positions_fields + positional_fields

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
            'Implied Total', 'O/U', 'Run DVOA', 'Pass DVOA', 'DVOA', 'WR1', 'WR2',
            'O-Line'
        ]
        green_to_red_rule = ColorScaleRule(start_type='min', start_color=red,
                                           mid_type='percentile', mid_value=50, mid_color=yellow,
                                           end_type='max', end_color=green)
        # bigger/positive = red, smaller/negative = green
        red_to_green_headers = [
            'Line', 'D-Line'
        ]
        red_to_green_rule = ColorScaleRule(start_type='min', start_color=green,
                                           mid_type='percentile', mid_value=50, mid_color=yellow,
                                           end_type='max', end_color=red)
        # color ranges
        for i in range(1, ws.max_column):

            if ws.cell(row=2, column=i).value in green_to_red_headers:
                column_letter = get_column_letter(i)
                # color range (green to red)
                cell_rng = "{0}{1}:{2}".format(column_letter, '3', ws.max_row)
                wb[title].conditional_formatting.add(cell_rng, green_to_red_rule)
            elif ws.cell(row=2, column=i).value in red_to_green_headers:
                column_letter = get_column_letter(i)
                # color range (red to green)
                cell_rng = "{0}{1}:{2}".format(column_letter, '3', ws.max_row)
                wb[title].conditional_formatting.add(cell_rng, red_to_green_rule)

        # set column widths
        column_widths = [8, 20, 10, 8, 8, 8, 8, 8, 8, 8, 8, 8]
        for i, column_width in enumerate(column_widths):
            ws.column_dimensions[get_column_letter(i + 1)].width = column_width


def main():
    fn = 'DKSalaries_week7_full.csv'
    dest_filename = 'sheet.xlsx'

    # create workbook/worksheet
    wb = Workbook()
    ws1 = wb.active
    ws1.title = 'DEL'

    # guess types (numbers, floats, etc)
    wb.guess_types = True

    with open(fn, 'r') as f:
        # read entire file into memory
        lines = f.readlines()

        for i, line in enumerate(lines):
            # append header to first worksheet, otherwise skip it
            if i == 0:
                continue

            fields = line.rstrip().split(',')
            position_tab(wb, fields, fields[0])
            # if fields[0] == 'RB':
            #     position_tab(wb, fields, 'RB')
            # elif fields[0] == 'WR':
            #     position_tab(wb, fields, 'WR')
            # else:
            #     # % salary DK
            #     salary_perc = "{0:.1%}".format(float(fields[5]) / 50000)
            #     salary = "${0}".format(fields[5])
            #
            #     my_fields = [fields[0], fields[2], salary, salary_perc, fields[6], fields[7], fields[8]]
            #
            #     # print fields to position-named worksheet
            #     print_position_ws(wb, fields[0], my_fields)
            #
            # # print fields to first worksheet
            # ws1.append(line.rstrip().split(','))

    # pull stats from lineups.com
    get_nfl_receptions(wb)
    get_nfl_targets(wb)
    get_nfl_snaps(wb)
    get_nfl_rush_atts(wb)
    # pull stats from footballoutsiders.com
    get_dvoa_rankings(wb)
    get_oline_rankings(wb)
    get_dline_rankings(wb)
    get_vegas_rg(wb)

    # color ranges
    style_ranges(wb)

    # save workbook (.xlsx file)
    wb.remove(ws1)  # remove blank worksheet
    wb.save(filename=dest_filename)


if __name__ == "__main__":
    main()
