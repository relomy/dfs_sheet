import csv
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
        lcell = row[0]
        rcell = row[-1]
        lcell.border = lcell.border + left
        rcell.border = rcell.border + right
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
        ENDPOINT = 'https://www.fantasypros.com/nfl/rankings/{}.php'.format(
            position.lower())
    else:
        ENDPOINT = 'https://www.fantasypros.com/nfl/rankings/ppr-{}.php'.format(
            position.lower())

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


def massage_name(name):
    # remove periods from name
    name = name.replace('.', '')
    # remove Jr. and III etc
    name = ' '.join(name.split(' ')[:2])
    # special fix for Juju Smith-Schuster
    name = name.replace('Juju', 'JuJu')

    return name


def get_lineups_player_stats():
    dictionary = {}
    dictionary['snaps'] = get_lineups_nfl_snaps()
    dictionary['targets'] = get_lineups_nfl_targets()
    dictionary['receptions'] = get_lineups_nfl_receptions()
    dictionary['rush_atts'] = get_lineups_nfl_rush_atts()
    return dictionary


def get_lineups_nfl_snaps():
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/snaps/2018/QB'
    fn = 'nfl_snaps.json'
    dir = 'sources'
    filename = path.join(dir, fn)

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = pull_data(filename, ENDPOINT)

    if data is None:
        raise Exception('Failed to pull data from API or file.')

    # create dictionary and set key to player's full name
    # dictionary = {}
    # for player in data['data']:
    #     dictionary[player['full_name']] = player

    # return dict comprehension for code above
    return {massage_name(x['full_name']): x for x in data['data']}


def get_lineups_nfl_targets():
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/targets/2018/OFF'
    fn = 'nfl_targets.json'
    dir = 'sources'
    filename = path.join(dir, fn)

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = pull_data(filename, ENDPOINT)

    if data is None:
        raise Exception('Failed to pull data from API or file.')

    # create dictionary and set key to player's full name
    return {massage_name(x['full_name']): x for x in data['data']}


def get_lineups_nfl_receptions():
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/receptions/2018/OFF'
    fn = 'nfl_receptions.json'
    dir = 'sources'
    filename = path.join(dir, fn)

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = pull_data(filename, ENDPOINT)

    if data is None:
        raise Exception('Failed to pull data from API or file.')

    # create dictionary and set key to player's full name
    return {massage_name(x['name']): x for x in data['data']}


def get_lineups_nfl_rush_atts():
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/rush/2018/OFF'
    fn = 'nfl_rush_atts.json'
    dir = 'sources'
    filename = path.join(dir, fn)

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = pull_data(filename, ENDPOINT)

    if data is None:
        raise Exception('Failed to pull data from API or file.')

    # create dictionary and set key to player's full name
    return {massage_name(x['name']): x for x in data['data']}


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
              'pass_yds', 'pass_tds', 'compl_perc', 'pass_td_per_att']
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
    dictionary = {}
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

        dictionary[team_abbv] = dict(zip(header, ls))
    return dictionary


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
    header = ['Time', 'Team', 'Opponent', 'Line', 'MoneyLine',
              'Over/Under', 'Projected Points', 'Projected Points Change']
    create_sheet_header(wb, title, header)

    # pull data
    soup = pull_soup_data(filename, ENDPOINT)

    # find script(s) in the html
    script = soup.findAll('script')

    js_vegas_data = script[11].string

    # replace two-letter abbvs
    js_vegas_data = js_vegas_data.replace('GBP', 'GB')
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

    dictionary = {}
    # iterate through json
    for matchup in vegas_json:
        dictionary[matchup['team']] = {
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
    return dictionary


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
        dict_dvoa_rankings_all = get_dvoa_recv_rankings(
            wb, table[1], title, dict_team_rankings)

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


def get_line_rankings(wb):
    dir = 'sources'
    # create empty dict to return
    dictionary = {}
    # dictionary['dl'] = {}
    for line in ['ol', 'dl']:
        ENDPOINT = "https://www.footballoutsiders.com/stats/{0}".format(line)
        fn = "html_{0}.html".format(line)
        filename = path.join(dir, fn)

        # create dict for run and pass stats
        dictionary[line] = {}
        dictionary[line]['run'] = {}
        dictionary[line]['pass'] = {}
        # print(dictionary[line]['run'])

        # pull data
        soup = pull_soup_data(filename, ENDPOINT)

        # find all tables (2) in the html
        table = soup.findAll('table')

        if table:
            # create worksheet
            if line == 'ol':
                title = 'OLINE'
            elif line == 'dl':
                title = 'DLINE'
            wb.create_sheet(title=title)

            # store table
            line_stats = table[0]
            # find header
            table_header = line_stats.find('thead')
            # there is one header row
            header_row = table_header.find('tr')
            # loop through header columns and append to worksheet
            header_cols = header_row.find_all('th')
            header = [ele.text.strip() for ele in header_cols]
            wb[title].append(header)

            # find the rest of the table header_rows
            rows = line_stats.find_all('tr')
            for row in rows:
                cols = row.find_all('td')
                cols = [ele.text.strip() for ele in cols]
                if cols:

                    # pop 'team_abbv' for dict key
                    run_key = cols.pop(1)

                    # pop 'team_abbv' for pass protection
                    pass_key = cols.pop(11)

                    run_key_names = ['rank', 'adj_line_yds', 'rb_yds', 'power_succ_perc', 'power_rank',
                                     'stuff_perc', 'stuff_rank', '2nd_lvl_yds', '2nd_lvl_rank',
                                     'open_field_yds', 'open_field_rank']

                    pass_key_names = ['rank', 'sacks', 'adj_sack_rate']
                    # print(key)

                    # map o-line to 'run' key
                    dictionary[line]['run'][run_key] = dict(zip(run_key_names, cols))
                    # map d-line to 'pass' key
                    dictionary[line]['pass'][pass_key] = dict(zip(pass_key_names, cols))
                    # example
                    # dictionary['ol']['run']['LAR']['adj_line_yds'] = 6.969
                    # dictionary['dl']['pass']['MIA']['adj_sack_rate'] = 4.5%

                    wb[title].append(cols)
    return dictionary


def qb_map(key):
    fo_qb_names = {
        'D.Brees': 'Drew Brees',
        'P.Mahomes': 'Patrick Mahomes',
        'J.Goff': 'Jared Goff',
        'P.Rivers': 'Phillip Rivers',
        'M.Ryan': 'Matt Ryan',
        'R.Fitzpatrick': 'Ryan Fitzpatrick',
        'A.Dalton': 'Andy Dalton',
        'J.Flacco': 'Joe Flacco',
        'A.Rodgers': 'Aaron Rodgers',
        'B.Roethlisberger': 'Ben Roethlisberger',
        'K.Cousins': 'Kirk Cousins',
        'T.Brady': 'Tom Brady',
        'D.Carr': 'Derek Carr',
        'M.Trubisky': 'Mitchell Trubisky',
        'D.Watson': 'Deshaun Watson',
        'C.Newton': 'Cam Newton',
        'C.Wentz': 'Carson Wentz',
        'R.Wilson': 'Russell Wilson',
        'J.Winston': 'Jameis Winston',
        'M.Stafford': 'Matthew Stafford',
        'S.Darnold': 'Sam Darnold',
        'A.Luck': 'Andrew Luck',
        'C.Keenum': 'Case Keenum',
        'C.J.Beathard': 'CJ Beathard',
        'A.Smith': 'Alex Smith',
        'J.Rosen': 'Josh Rosen',
        'B.Bortles': 'Blake Bortles',
        'E.Manning': 'Eli Manning',
        'J.Garoppolo': 'Jimmy Garoppolo',
        'R.Tannehill': 'Ryan Tannehill',
        'D.Prescott': 'Dak Prescott',
        'M.Mariota': 'Marcus Mariota',
        'B.Mayfield': 'Baker Mayfield',
        'T.Taylor': 'Tyrod Taylor',
        'J.Allen': 'Josh Allen',
        'B.Osweiler': 'Brock Osweiler',
        'B.Gabbert': 'Blaine Gabbert',
        'C.Kessley': 'Cody Kessler',
        'D.Enderson': 'Derek Anderson',
        'N.Foles': 'Nick Foles',
        'S.Bradford': 'Sam Bradford',
        'N.Peterman': 'Nathan Peterman',
        'T.Taylor': 'Tyrod Taylor',
        'L.Jackson': 'Lamar Jackson'
    }
    return fo_qb_names.get(key, None)


def get_qb_stats_FO(wb):
    ENDPOINT = 'https://www.footballoutsiders.com/stats/qb'
    fn = 'html_qb.html'
    dir = 'sources'
    filename = path.join(dir, fn)

    # pull data
    soup = pull_soup_data(filename, ENDPOINT)

    # find all tables (3) in the html
    table = soup.findAll('table')

    if table:
        # create worksheet
        title = 'QB_STATS'
        wb.create_sheet(title=title)

        dictionary = {}
        for i, t in enumerate(table):
            # find header
            table_header = t.find('thead')
            # there is one header row
            header_row = table_header.find('tr')
            # loop through header columns and append to worksheet
            header_cols = header_row.find_all('th')
            header = [ele.text.strip() for ele in header_cols]
            wb[title].append(header)

            # find the rest of the table header_rows
            rows = t.find_all('tr')
            for row in rows:
                cols = row.find_all('td')
                cols = [ele.text.strip() for ele in cols]
                if cols:
                    # pop 'name' for dict key
                    key = cols.pop(0)

                    # i only create this list because Lamar Jackson has no pass_dyar but he has rushing stats
                    main_fields = ['team', 'pass_dyar', 'dyar_rank', 'yar', 'yar_rank', 'pass_dvoa', 'pass_dvoa_rank', 'voa', 'qbr',
                                   'qbr_rank', 'pass_atts', 'pass_yds', 'eyds', 'tds', 'fk', 'fl', 'int', 'c_perc', 'dpi', 'alex']
                    if i == 0:
                        key_names = main_fields
                    elif i == 1:
                        key_names = ['team', 'pass_dyar', 'pass_dvoa', 'pass_dvoa_rank', 'voa', 'qbr',
                                     'qbr_rank', 'pass_atts', 'pass_yds', 'eyds', 'tds', 'fk', 'fl', 'int', 'c_perc', 'dpi', 'alex']
                    elif i == 2:
                        key_names = ['team', 'rush_dyar', 'rush_dyar_rank', 'rush_yar', 'rush_yar_rank', 'rush_dvoa', 'rush_dvoa_rank',
                                     'rush_voa', 'rush_atts', 'rush_yds', 'rush_eyds', 'rush_tds', 'fumbles']

                    # print(key)
                    # map key_names to cols
                    player_name = qb_map(key)
                    # create dictionary if it does not exist
                    if player_name not in dictionary:
                        dictionary[player_name] = dict.fromkeys(main_fields, None)
                    dictionary[player_name].update(dict(zip(key_names, cols)))

                    wb[title].append(cols)
    return dictionary


def find_name_in_ecr(ecr_pos_list, name):
    for item in ecr_pos_list:
        if any(name in s for s in item):
            # print("Found {}!".format(name))
            # rank, wsis, dumb_name, matchup, best, worse, avg, std_dev = item
            return item
    return False


def read_fantasy_draft_csv(filename):
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

    with open(filename, 'r') as f:
        reader = csv.reader(f)

        # store header row (and strip extra spaces)
        headers = [header.lower().strip() for header in next(reader)]
        headers.append('salary_perc')

        # fill dictionary to return
        dictionary = {}
        for row in reader:
            if row[0] == 'DST':
                # map full team name to team abbv
                row[1] = team_map.get(row[1], None)
            # remove periods from name
            row[1] = row[1].replace('.', '')
            # remove Jr. and III etc
            row[1] = ' '.join(row[1].split(' ')[:2])
            # store salary without $ or ,
            row[5] = row[5][1:].replace(',', '')
            # calculate salary percentage
            salary_perc = "{0:0.1%}".format(float(row[5]) / 100000)
            row.append(salary_perc)
            dictionary[row[1]] = {key: value for key, value in zip(headers, row)}
        return dictionary


def write_position_to_sheet(wb, player):
    if player.position not in wb.sheetnames:
        wb.create_sheet(title=player.position)
        wb[player.position].append(player.get_writable_header())

    ws = wb[player.position]

    ws.append(player.get_writable_row())


def main():
    fn = 'DKSalaries_week8_full.csv'
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

    fdraft_csv = 'FDraft_week8_full.csv'
    if path.exists(fdraft_csv):
        fdraft_dict = read_fantasy_draft_csv(fdraft_csv)
    else:
        fdraft_dict = None

    # vegas lines from rotogrinders.com
    vegas_dict = get_vegas_rg(wb)
    # get snaps, targets, receptions, rush attempts from lineups.com
    stats_dict = get_lineups_player_stats()
    # defense stats from lineups.com
    def_dict = get_nfl_def_stats(wb)
    # DVOA rankings from footballoutsiders.com
    dvoa_dict = get_dvoa_rankings(wb)
    # OL/DL rankings from footballoutsiders.com
    line_dict = get_line_rankings(wb)
    # QB rankings from footballoutsiders.com
    qb_dict = get_qb_stats_FO(wb)

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

            # store each variable from list
            position, name_id, name, id, roster_pos, salary, game_info, team_abbv, average_ppg = fields

            # 'fix' name to remove extra stuff like Jr or III (Todd Gurley II for example)
            name = ' '.join(name.split(' ')[:2])
            # also remove periods (T.J. Yeldon for example)
            name = name.replace('.', '')

            if position == 'DST':
                name = fields[7]

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
                p = Player(name, position, team_abbv, salary, game_info,
                           average_ppg, ecr_matchup, ecr_rank)

                if fdraft_dict:
                    p.set_fdraft_fields(fdraft_dict[name]['salary'],
                                        fdraft_dict[name]['salary_perc'])
                # local variable for dicts
                dvoa_opponent = dvoa_dict[p.opponent]
                # local variable for player's team
                vegas_player_team = vegas_dict[team_abbv]
                # set vegas fields based on team abbv (key)
                p.set_vegas_fields(
                    vegas_player_team['overunder'], vegas_player_team['line'], vegas_player_team['projected'])

                if position == 'QB':
                    qb = QB(p)
                    qb.set_sack_fields(line_dict['dl']['pass'][team_abbv]['adj_sack_rate'],
                                       line_dict['dl']['pass'][p.opponent]['adj_sack_rate'])
                    qb.set_season_fields(
                        qb_dict[name]['rush_yds'], qb_dict[name]['pass_dyar'], qb_dict[name]['qbr'])
                    qb.set_matchup_fields(
                        def_dict[p.opponent]['pass_yd_per_att'], def_dict[p.opponent]['compl_perc'], def_dict[p.opponent]['pass_td_per_att'])

                    player_list.append(qb)
                elif position == 'RB':
                    rb = RB(p)
                    # set position-specific dvoa fields
                    rb.set_dvoa_fields(
                        dvoa_opponent['rush_def_rank'], dvoa_opponent['rb_rank'])
                    # set oline/opponent dline stats for adjusted line yards
                    rb.set_line_fields(line_dict['ol']['run'][team_abbv]['adj_line_yds'],
                                       line_dict['dl']['run'][p.opponent]['adj_line_yds'])
                    if name in stats_dict['snaps']:
                        rb.set_season_fields(stats_dict['snaps'][name]['average'],
                                             stats_dict['rush_atts'][name]['average'],
                                             stats_dict['targets'][name]['average'])

                        # store lists in Player object
                        rb.snap_percentage_by_week = stats_dict['snaps'][name]['snap_percentage_by_week']
                        rb.rush_atts_weeks = stats_dict['rush_atts'][name]['weeks']
                        rb.targets_weeks = stats_dict['targets'][name]['weeks']
                        print(rb)
                        # print("snaps list: {}".format(rb.snap_percentage_by_week))
                        # print("last_week_snaps: {}".format(rb.last_week_snaps()))
                        print("rush dict: {}".format(rb.rush_atts_weeks))
                        print("last_week_rush: {}".format(rb.last_week_rush_atts()))

                        print("targets dict: {}".format(rb.targets_weeks))
                        print("targets list: {}".format(rb.targets_weeks[-1]))
                        print("last_week_targets: {}".format(rb.last_week_targets()))
                        exit()
                        rb.set_last_week_fields('x', 'x', 'x')
                    else:
                        print("Could find no SNAPS information on {}".format(name))
                    player_list.append(rb)
                elif position == 'WR':
                    wr = WR(p)
                    # set position-specific dvoa fields
                    wr.set_dvoa_fields(dvoa_opponent['pass_def_rank'],
                                       dvoa_opponent['rb_rank'], dvoa_opponent['rb_rank'])
                    # if player is not in snaps, he likely has no other information either
                    if name in stats_dict['snaps']:
                        wr.set_season_fields(stats_dict['snaps'][name]['average'],
                                             stats_dict['targets'][name]['average'],
                                             stats_dict['receptions'][name]['average'])
                        wr.set_last_week_fields('x', 'x', 'x')
                    else:
                        print("Could find no SNAPS information on {}".format(name))
                    player_list.append(wr)
                elif position == 'TE':
                    te = TE(p)
                    # set position-specific dvoa fields
                    te.set_dvoa_fields(
                        dvoa_opponent['pass_def_rank'], dvoa_opponent['te_rank'])
                    if name in stats_dict['snaps']:
                        te.set_season_fields(stats_dict['snaps'][name]['average'],
                                             stats_dict['targets'][name]['average'],
                                             stats_dict['receptions'][name]['average'])
                        te.set_last_week_fields('x', 'x', 'x')
                    else:
                        print("Could find no SNAPS information on {}".format(name))
                    player_list.append(te)
                elif position == 'DST':
                    dst = DST(p)
                    player_list.append(dst)
            # position_tab(wb, fields, fields[0])

    # for k, v in player_dict.items():
    #     print("k: {}".format(k))
    #     print("v: {}".format(v))
    for i, player in enumerate(player_list):
        # if player.position == 'QB':
            # write_position_to_sheet(wb, player)
        write_position_to_sheet(wb, player)
        # print(player)
        # print(player.fdraft_salary)
        # print(player.fdraft_salary_perc)

        # print("run_dvoa: {} pass_dvoa: {}".format(rb.run_dvoa, rb.rb_pass_dvoa))
        # print("[{}] ou: {} line: {} proj: {}".format(rb.team_abbv, rb.overunder, rb.line, rb.projected))

    # save workbook (.xlsx file)
    wb.remove(ws1)  # remove blank worksheet
    wb.save(filename=dest_filename)


if __name__ == "__main__":
    main()
