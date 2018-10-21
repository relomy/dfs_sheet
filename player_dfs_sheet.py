import json
import re
import requests
from os import path
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font, colors

from player import Player, QB, RB, WR, TE, DST


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


def get_nfl_def_stats(wb):
    # https://www.lineups.com/nfl/teams/stats/defense-stats
    # get passing yds/att
    # td / att (td %)
    # att / completion (compl %)
    """Retrieve receptions from lineups.com API."""
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/teams/stats/defense-stats/current'
    fn = 'nfl_def_stats.json'
    dir = 'sources'
    filename = path.join(dir, fn)

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = pull_data(filename, ENDPOINT)

    # we just want player data
    player_data = data['data']

    # create worksheet
    title = 'DEF_STATS'
    header = ['team abbv', 'team', 'pass_att', 'pass_yd_per_att', 'pass_compls', 'pass_yd_per_compl',
              'pass_yds', 'pass_tds', 'pass_td_per_att', 'compl_perc']
    create_sheet_header(wb, title, header)

    team_map = {
        'Atlanta Falcons': 'ATL',
        'Indianapolis Colts': 'IND',
        'San Francisco 49ers': 'SF',
        'Oakland Raiders': 'OAK',
        'Tampa Bay Buccaneers': 'TB',
        'Kansas City Chiefs': 'KC',
        'New York Giants': 'NYG',
        'Cincinnati Bengals': 'CIN',
        'Pittsburgh Steelers': 'PIT',
        'Denver Broncos': 'DEN',
        'Cleveland Browns': 'CLE',
        'New England Patriots': 'NE',
        'Minnesota Vikings': 'MIN',
        'Miami Dolphins': 'MIA',
        'Green Bay Packers': 'GB',
        'Los Angeles Chargers': 'LAC',
        'New Orleans Saints': 'NO',
        'New York Jets': 'NYJ',
        'Arizona Cardinals': 'ARI',
        'Buffalo Bills': 'BUF',
        'Houston Texans': 'HOU',
        'Detroit Lions': 'DET',
        'Jacksonville Jaguars': 'JAX',
        'Los Angeles Rams': 'LAR',
        'Seattle Seahawks': 'SEA',
        'Philadelphia Eagles': 'PHI',
        'Carolina Panthers': 'CAR',
        'Tennessee Titans': 'TEN',
        'Washington Redskins': 'WAS',
        'Dallas Cowboys': 'DAL',
        'Chicago Bears': 'CHI',
        'Baltimore Ravens': 'BAL'
    }

    for d in player_data:
        # TODO rushing_attempt_percentage_by_week
        team = d['team']
        team_abbv = team_map[team]
        pass_att = d['passing_attempts']
        pass_yd_per_att = d['passing_yards_per_attempt']
        pass_compls = d['passing_completions']
        pass_yd_per_compl = d['passing_yards_per_completion']
        pass_yds = d['passing_yards']
        pass_tds = d['passing_touchdowns']

        # personal
        pass_td_per_att = "{0:.4f}".format(pass_tds / pass_att)
        compl_perc = "{0:.4f}".format(pass_compls / pass_att)

        # remove '.' from name
        # name = name.replace('.', '')

        ls = [team_abbv, team, pass_att, pass_yd_per_att, pass_compls, pass_yd_per_compl,
              pass_yds, pass_tds, compl_perc, pass_td_per_att]

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

    # pull data
    soup = pull_soup_data(filename, ENDPOINT)

    # find script(s) in the html
    script = soup.findAll('script')

    js_vegas_data = script[11].string

    # replace dumb names
    js_vegas_data = js_vegas_data.replace('JAC', 'JAX')
    js_vegas_data = js_vegas_data.replace('KCC', 'KC')
    js_vegas_data = js_vegas_data.replace('NEP', 'NE')
    js_vegas_data = js_vegas_data.replace('NOS', 'NO')
    js_vegas_data = js_vegas_data.replace('SFO', 'SF')
    js_vegas_data = js_vegas_data.replace('TBB', 'TB')

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

    # make blank NFL key
    return_dict['NFL'] = {}
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
            dict_team_rankings[key].update(dict(zip(key_names, cols)))

            wb[title].append(cols)

    return dict_team_rankings


def find_name_in_ecr(ecr_pos_list, name):
    for item in ecr_pos_list:
        if any(name in s for s in item):
            # print("Found {}!".format(name))
            # rank, wsis, dumb_name, matchup, best, worse, avg, std_dev = item
            return item

    return False


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
    for position in ['QB', 'RB', 'WR', 'TE', 'DST']:
        ecr_pos_dict[position] = fpros_ecr(wb, position)

    # pull data
    vegas_dict = get_vegas_rg(wb)
    dvoa_dict = get_dvoa_rankings(wb)

    # print(dvoa_dict['CHI'])

    # create list for players
    player_list = []

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

            position, name_id, name, id, roster_pos, salary, game_info, team_abbv, average_ppg = fields

            # print("opp: {} opp_excel: {}".format(opp, opp_excel))
            # if player does not exist, skip
            # for item in ecr_pos_dict[position]:
            #     if any(name in s for s in item):
            #         print("Found {}!".format(name))
            ecr_item = find_name_in_ecr(ecr_pos_dict[position], name)
            if ecr_item:
                # ecr_rank, ecr_wsis, ecr_dumb_name, ecr_matchup, ecr_best, ecr_worse, ecr_avg, ecr_std_dev = ecr_item
                ecr_rank = ecr_item[0]
                ecr_matchup = ecr_item[3]
                # create Player class for position
                p = Player(name, position, team_abbv, salary, game_info, average_ppg, ecr_matchup, ecr_rank)
                # set vegas fields based on team abbv (key)
                p.set_vegas_fields(vegas_dict[team_abbv]['overunder'], vegas_dict[team_abbv]['line'], vegas_dict[team_abbv]['projected'])

                # local variable for dvoa_dict
                dvoa_opponent = dvoa_dict[p.opponent]

                if position == 'QB':
                    qb = QB(p)
                    # set vegas fields based on team abbv (key)
                    player_list.append(qb)
                elif position == 'RB':
                    rb = RB(p)
                    # set position-specific dvoa fields
                    rb.set_dvoa_fields(dvoa_opponent['rush_def_rank'], dvoa_opponent['rb_rank'])
                    player_list.append(rb)
                elif position == 'WR':
                    wr = WR(p)
                    # set position-specific dvoa fields
                    wr.set_dvoa_fields(dvoa_opponent['pass_def_rank'], dvoa_opponent['rb_rank'], dvoa_opponent['rb_rank'])
                    player_list.append(wr)
                elif position == 'TE':
                    te = TE(p)
                    # set position-specific dvoa fields
                    te.set_dvoa_fields(dvoa_opponent['pass_def_rank'], dvoa_opponent['te_rank'])
                    player_list.append(te)
                elif position == 'DST':
                    dst = DST(p)
                    player_list.append(dst)
            # position_tab(wb, fields, fields[0])

    # for k, v in player_dict.items():
    #     print("k: {}".format(k))
    #     print("v: {}".format(v))
    for i, player in enumerate(player_list):
        if player.position == 'TE':
            print(player)
            print(player.te_rank)
            print(player.pass_def_rank)
        # print("run_dvoa: {} pass_dvoa: {}".format(rb.run_dvoa, rb.rb_pass_dvoa))
        # print("[{}] ou: {} line: {} proj: {}".format(rb.team_abbv, rb.overunder, rb.line, rb.projected))

    # save workbook (.xlsx file)
    wb.remove(ws1)  # remove blank worksheet
    wb.save(filename=dest_filename)


if __name__ == "__main__":
    main()
