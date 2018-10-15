import json
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


def get_nfl_targets(wb):
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/targets/2018/RB'
    filename = 'nfl_targets.json'

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = ''
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

    player_data = data['data']

    # create worksheet
    wb.create_sheet(title="TARGETS")
    header = ['name', 'position', 'rating', 'team', 'week1', 'week2', 'week3', 'week4', 'week5', 'week6',
              'week7', 'week8', 'week9', 'week10', 'week11', 'week12', 'week13', 'week14',
              'week15', 'week16', 'targets', 'average', 'recv touchdowns']
    wb["TARGETS"].append(header)

    for d in player_data:
        name = d['full_name']
        position = d['position']
        rating = d['lineups_rating']
        team = d['team']
        targets = d['total']
        weeks = d['weeks']  # dict
        average = d['average']
        recv_touchdowns = d['receiving_touchdowns']
        catch_percentage = d['catch_percentage']
        season_target_percent = d['season_target_percent']

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
        pre_weeks = [name, position, rating, team]
        post_weeks = [targets, average, recv_touchdowns]
        ls = pre_weeks + all_weeks + post_weeks

        # insert all_weeks list into ls
        # ls = [name, position, rating, team, receptions, average, touchdowns]
        # print("trying to insert: ls[2:{}]".format(len(all_weeks)))
        # ls[4:len(all_weeks)-1] = all_weeks
        # print(ls)

        wb["TARGETS"].append(ls)


def get_nfl_receptions(wb):
    ENDPOINT = 'https://api.lineups.com/nfl/fetch/receptions/2018/RB'
    filename = 'nfl_receptions.json'

    # if file doesn't exist, let's pull it. otherwise - use the file.
    data = ''
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

    player_data = data['data']

    # create worksheet
    wb.create_sheet(title="RECEPTIONS")
    header = ['name', 'position', 'rating', 'team', 'week1', 'week2', 'week3', 'week4', 'week5', 'week6',
              'week7', 'week8', 'week9', 'week10', 'week11', 'week12', 'week13', 'week14',
              'week15', 'week16', 'receptions', 'average', 'touchdowns']
    wb["RECEPTIONS"].append(header)

    for d in player_data:
        #{'receptions': 0.0, 'id': 20977, 'lineups_rating': None, 'total': 0.0, 'team': 'DET', 'profile_url': '/nfl/player-stats/ameer-abdullah', 'fantasy_position_depth_order': 4, 'average': 0.0, 'position': 'RB', 'touchdowns': 0.0, 'name': 'Ameer Abdullah', 'team_depth_chart_route': '/nfl/depth-charts/detroit-lions', 'weeks': {'1': None, '2': None, '3': None, '4': None, '5': None}}

        name = d['name']
        position = d['position']
        rating = d['lineups_rating']
        team = d['team']
        receptions = d['receptions']
        weeks = d['weeks']  # dict
        average = d['average']
        touchdowns = d['touchdowns']

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
        pre_weeks = [name, position, rating, team]
        post_weeks = [receptions, average, touchdowns]
        ls = pre_weeks + all_weeks + post_weeks

        # insert all_weeks list into ls
        # ls = [name, position, rating, team, receptions, average, touchdowns]
        # print("trying to insert: ls[2:{}]".format(len(all_weeks)))
        # ls[4:len(all_weeks)-1] = all_weeks
        # print(ls)

        wb["RECEPTIONS"].append(ls)


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

    # pull stats
    get_nfl_receptions(wb)
    get_nfl_targets(wb)

    # save workbook (.xlsx file)
    wb.save(filename=dest_filename)


if __name__ == "__main__":
    main()
