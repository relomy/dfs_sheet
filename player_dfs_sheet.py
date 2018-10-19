import json
import re
import requests
from os import path
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font, colors

from player import QB


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


def create_sheet_header(wb, title, header):
    wb.create_sheet(title=title)
    wb[title].append(header)


def pull_soup_data(filename, ENDPOINT):
    """Either pull file from html or from file."""
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

    return soup


def fpros_ecr(wb, position):
    if position == 'QB' or position == 'DST':
        ENDPOINT = 'https://www.fantasypros.com/nfl/rankings/{}.php'.format(position.lower())
    else:
        ENDPOINT = 'https://www.fantasypros.com/nfl/rankings/ppr-{}.php'.format(position.lower())

    fn = 'ecr_{}.html'.format(position)
    dir = 'sources'
    filename = path.join(dir, fn)

    # pull data
    soup = pull_soup_data(filename, ENDPOINT)

    # find all tables (2) in the html
    table = soup.find('table', id='rank-data')

    if table:
        # create worksheet
        title = '{0}_ECR'.format(position)
        wb.create_sheet(title=title)
        ls = []

        # # find header
        table_header = table.find('thead')
        # there is one header row
        header_row = table_header.find('tr')
        # loop through header columns and append to worksheet
        header_cols = header_row.find_all('th')
        header = [ele.text.strip() for ele in header_cols]
        wb[title].append(header)
        ls.append(header)

        # ignore notes
        header = header[:-1]

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
                ls.append(new_cols)
                wb[title].append(new_cols)

        # return dict(zip(header, new_cols))
        return ls


def get_vegas_rg(wb):
    ENDPOINT = 'https://rotogrinders.com/schedules/nfl'

    fn = 'vegas_script.html'
    dir = 'sources'
    filename = path.join(dir, fn)

    # create worksheet
    title = 'VEGAS'
    header = ['Time', 'Team', 'Opponent', 'Line', 'MoneyLine', 'Over/Under', 'Projected Points', 'Projected Points Change']
    create_sheet_header(wb, title, header)

    # pull data
    soup = pull_soup_data(filename, ENDPOINT)

    # find script(s) in the html
    script = soup.findAll('script')

    js_vegas_data = script[11].string

    # replace dumb names
    js_vegas_data = js_vegas_data.replace('KCC', 'KC')
    js_vegas_data = js_vegas_data.replace('JAC', 'JAX')
    js_vegas_data = js_vegas_data.replace('TBB', 'TB')
    js_vegas_data = js_vegas_data.replace('NEP', 'NE')
    js_vegas_data = js_vegas_data.replace('NOS', 'NO')

    # pull json object from data variable
    pattern = re.compile(r'data = (.*);')
    json_str = pattern.search(js_vegas_data).group(1)
    vegas_json = json.loads(json_str)

    return_dict = {}
    # iterate through json
    for matchup in vegas_json:
        return_dict[matchup['team']] = {
            'display_time': matchup['time']['display'],
            'opponent': matchup['opponent'],
            'line': matchup['line'],
            'moneyline': matchup['moneyline'],
            'overunder': matchup['overunder'],
            'projected': matchup['projected'],
            'projectedchange': matchup['projectedchange']['value']
        }
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
    return return_dict


def get_dvoa_rankings(wb):
    ENDPOINT = 'https://www.footballoutsiders.com/stats/teamdef'
    fn = 'html_defense.html'
    dir = 'sources'
    filename = path.join(dir, fn)

    # pull data
    soup = pull_soup_data(filename, ENDPOINT)

    # find all tables (3) in the html
    table = soup.findAll('table')

    if table:
        # create worksheet
        title = 'TEAMDEF'
        wb.create_sheet(title=title)

        dict_team_rankings = get_dvoa_team_rankings(wb, table[0], title)
        # separate function for second table
        dict_dvoa_rankings_all = get_dvoa_recv_rankings(wb, table[1], title, dict_team_rankings)

        return dict_dvoa_rankings_all


def get_dvoa_team_rankings(wb, soup_table, title):
    defense_stats = soup_table

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
    # na = non-adjusted

    return_dict = {}
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]

        if cols:
            # pop 'team_abbv' for dict key
            key = cols.pop(1)

            key_names = ['row_num', 'defense_dvoa', 'last_week',
                         'defense_dave', 'total_def_rank', 'pass_def', 'pass_def_rank', 'rush_def',
                         'rush_def_rank', 'na_total', 'na_pass', 'na_rush', 'var', 'sched', 'rank']
            # print(key)

            # map key_names to cols

            return_dict[key] = dict(zip(key_names, cols))

            # print columns to worksheet
            wb[title].append(cols)
            # return_dict = {
            #     'row_num': cols[0],
            #     'team_abbv': cols[1],
            #     'defense_dvoa': cols[2],
            #     'last_week': cols[3],
            #     'defense_dave': cols[4],
            #     'total_def_rank': cols[5],
            #     'pass_def': cols[6],
            #     'pass_def_rank': cols[7],
            #     'rush_def': cols[8],
            #     'rush_def_rank': cols[9],
            #     # na = non-adjusted
            #     'na_total': cols[10],
            #     'na_pass': cols[11],
            #     'na_rush': cols[12],
            #     'var': cols[13],
            #     'sched': cols[14],
            #     'rank': cols[15],
            # }

    return return_dict


def get_dvoa_recv_rankings(wb, soup_table, title, dict_team_rankings):
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
            # pop 'team_abbv' for dict key
            key = cols.pop(1)

            key_names = ['rk', 'wr1_dvoa', 'wr1_rank', 'wr1_pa_g', 'wr1_yd_g',
                         'wr2_dvoa', 'wr2_rank', 'wr2_pa_g', 'wr2_yd_g',
                         'wro_dvoa', 'wro_rank', 'wro_pa_g', 'wro_yd_g',
                         'te_dvoa', 'te_rank', 'te_pa_g', 'te_yd_g',
                         'rb_dvoa', 'rb_rank', 'rb_pa_g', 'rb_yd_g']
            # print(key)

            # map key_names to cols
            for col in cols:
                print("key: {} col: {}".format(key, col))
                dict_team_rankings[key] = dict(zip(key_names, cols))

            # return_dict_2 = {
            #     'rk': cols[0],
            #     # 'team_abbv': cols[1],
            #     # vs. WR1
            #     'wr1_dvoa': cols[2],
            #     'wr1_rank': cols[3],
            #     'wr1_pa_g': cols[4],
            #     'wr1_yd_g': cols[5],
            #     # vs. WR2
            #     'wr2_dvoa': cols[6],
            #     'wr2_rank': cols[7],
            #     'wr2_pa_g': cols[8],
            #     'wr2_yd_g': cols[9],
            #     # vs. Other WR
            #     'wro_dvoa': cols[10],
            #     'wro_rank': cols[11],
            #     'wro_pa_g': cols[12],
            #     'wro_yd_g': cols[13],
            #     # vs. TE
            #     'te_dvoa': cols[14],
            #     'te_rank': cols[15],
            #     'te_pa_g': cols[16],
            #     'te_yd_g': cols[17],
            #     # vs. RB
            #     'rb_dvoa': cols[18],
            #     'rb_rank': cols[19],
            #     'rb_pa_g': cols[20],
            #     'rb_yd_g': cols[21],
            # }
            # print(return_dict_2)
            wb[title].append(cols)

        return dict_team_rankings


def find_name_in_ecr(ecr_pos_list, name):
    for item in ecr_pos_list:
        if any(name in s for s in item):
            # print("Found {}!".format(name))
            # rank, wsis, dumb_name, matchup, best, worse, avg, std_dev = item
            return item

    return False


def get_opponent_matchup(game_info, team_abbv):
    print(game_info)
    print(game_info.split(' ', 1))
    home_team, away_team = game_info.split(' ', 1)[0].split('@')
    if team_abbv == home_team:
        opp = away_team
        opp_excel = "vs. {}".format(away_team)
    else:
        opp = home_team
        opp_excel = "at {}".format(home_team)

    return opp, opp_excel


def main():
    fn = 'DKSalaries_week7_full.csv'
    dest_filename = 'player_sheet.xlsx'

    # create workbook/worksheet
    wb = Workbook()
    ws1 = wb.active
    ws1.title = 'DEL'

    # guess types (numbers, floats, etc)
    wb.guess_types = True

    # dict  of players (key = DFS player name)
    # player_dict = {}

    # make sources dir if it does not exist
    # directory = 'sources'
    # if not path.exists(directory):
    #     makedirs(directory)

    # pull positional stats from fantasypros.com
    # for position in ['QB', 'RB', 'WR', 'TE', 'DST']:

    ecr_pos_dict = {}
    # for position in ['QB', 'RB', 'WR', 'TE', 'DST']:
    for position in ['QB']:
        ecr_pos_dict[position] = fpros_ecr(wb, position)

    vegas_dict = get_vegas_rg(wb)

    dvoa_dict = get_dvoa_rankings(wb)

    qb_list = []

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

            if position in ['RB', 'WR', 'TE', 'DST']:
                continue

            position, name_id, name, id, roster_pos, salary, game_info, team_abbv, average_ppg = fields

            opp, opp_excel = get_opponent_matchup(game_info, team_abbv)
            print("opp: {} opp_excel: {}".format(opp, opp_excel))
            # if player does not exist, skip
            # for item in ecr_pos_dict[position]:
            #     if any(name in s for s in item):
            #         print("Found {}!".format(name))
            ecr_item = find_name_in_ecr(ecr_pos_dict[position], name)
            if ecr_item:
                ecr_rank, ecr_wsis, ecr_dumb_name, ecr_matchup, ecr_best, ecr_worse, ecr_avg, ecr_std_dev = ecr_item
                # create Player subclass for position
                if position == 'QB':
                    qb = QB(position, name, team_abbv, salary, game_info, average_ppg)
                    qb.set_ecr_fields(ecr_matchup, ecr_rank)
                    # set vegas fields based on team abbv (key)
                    qb.set_vegas_fields(vegas_dict[team_abbv]['overunder'], vegas_dict[team_abbv]['line'], vegas_dict[team_abbv]['projected'])
                    qb_list.append(qb)


            # position_tab(wb, fields, fields[0])

    # for k, v in player_dict.items():
    #     print("k: {}".format(k))
    #     print("v: {}".format(v))
    # for qb in qb_list:
    #     print(qb)
    #     print("rank: {} matchup: {}".format(qb.rank, qb.matchup))
    #     print("[{}] ou: {} line: {} proj: {}".format(qb.abbv, qb.overunder, qb.line, qb.projected))


    # save workbook (.xlsx file)
    wb.remove(ws1)  # remove blank worksheet
    wb.save(filename=dest_filename)


if __name__ == "__main__":
    main()
