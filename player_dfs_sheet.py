import requests
from os import path
from bs4 import BeautifulSoup
from openpyxl import Workbook

from player import Player


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


def check_name_in_ecr(wb, position, name):
    # get ECR sheet
    ecr_ws = wb[position + '_ECR']
    search_col = 'C'

    # search ECR sheet for guy
    return bool_found_player_in_ecr_tab(ecr_ws[search_col], name)


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

def main():
    fn = 'DKSalaries_week7_full.csv'
    dest_filename = 'sheet.xlsx'

    # create workbook/worksheet
    wb = Workbook()
    ws1 = wb.active
    ws1.title = 'DEL'

    # guess types (numbers, floats, etc)
    wb.guess_types = True

    # dict  of players (key = DFS player name)
    player_dict = {}
    # make sources dir if it does not exist
    # directory = 'sources'
    # if not path.exists(directory):
    #     makedirs(directory)

    # pull positional stats from fantasypros.com
    # for position in ['QB', 'RB', 'WR', 'TE', 'DST']:

    all_pos_dict = {}
    # for position in ['QB', 'RB', 'WR', 'TE', 'DST']:
    for position in ['QB']:
        all_pos_dict[position] = fpros_ecr(wb, position)

    # for k, v in all_pos_dict.items():
        # print("k: {} v: {}".format(k, v))
        # for l in v:
        #     print(l)
    # print(all_pos_dict)

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

            # if player does not exist, skip
            for item in all_pos_dict[position]:
                if any(name in s for s in item):
                    # print("Found {}!".format(name))
                    rank, wsis, dumb_name, matchup, best, worse, avg, std_dev = item
                    player_dict[name] = Player(name, matchup, rank)
                    break
                # exit()

            # if check_name_in_ecr(wb, position, name) is False:
                # print("Could not find {} [{}]".format(name, position))
                # continue

            # position_tab(wb, fields, fields[0])

    for k, v in player_dict.items():
        print("k: {}".format(k))
        print("v: {}".format(v))

    # save workbook (.xlsx file)
    wb.remove(ws1)  # remove blank worksheet
    wb.save(filename=dest_filename)


if __name__ == "__main__":
    main()
