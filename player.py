"""Object to hold various stats per player."""


class Player:
    """Creates Player object."""

    def __init__(self, player_name, position, team_abbv, salary, game_info, average_ppg, matchup, rank):
        self.name = player_name

        # from DK salary CSV
        self.name = player_name
        self.position = position
        self.team_abbv = team_abbv
        self.salary = salary
        self.game_info = game_info
        self.average_ppg = average_ppg

        # fantasy draft salary CSV
        self.fdraft_salary = None
        self.fdraft_salary_perc = None

        # fantasy pros ECR
        self.matchup = matchup
        self.rank = rank

        self.opponent, self.opp_excel, self.home_team = self.get_opponent_matchup(
            game_info, team_abbv)

        # calculate salary percent
        self.salary_percent = "{0:.1%}".format(float(salary) / 50000)

        # ECR
        self.matchup = matchup
        self.rank = rank

        # vegas
        self.overunder = None
        self.line = None
        self.projected = None

        # rankings
        # self.average_ppg = average_ppg
        self.ecr = None
        self.ecr_data = None
        self.plus_minus = None
        self.salary_rank = None

    def assign(self, p):
        """Class method to assign class variables (for use in subclasses)."""
        self.name = p.name
        self.position = p.position
        self.team_abbv = p.team_abbv
        self.salary = p.salary
        self.game_info = p.game_info
        self.average_ppg = p.average_ppg
        self.matchup = p.matchup
        self.rank = p.rank

        # calculated field
        self.salary_percent = p.salary_percent

        # vegas
        self.overunder = p.overunder
        self.line = p.line
        self.projected = p.projected

        # fantasy draft salary CSV
        self.fdraft_salary = p.fdraft_salary
        self.fdraft_salary_perc = p.fdraft_salary_perc

        # variables from get_opponent_matchup()
        self.opponent = p.opponent
        self.opp_excel = p.opp_excel
        self.home_team = p.home_team

    def __repr__(self):
        return "Player({}, {})".format(self.name, self.rank)

    def set_fdraft_fields(self, fdraft_salary, fdraft_salary_perc):
        self.fdraft_salary = fdraft_salary
        self.fdraft_salary_perc = fdraft_salary_perc

    def set_vegas_fields(self, overunder, line, projected):
        self.overunder = overunder
        self.line = line
        self.projected = projected

    def get_opponent_matchup(self, game_info, team_abbv):
        home_team, away_team = game_info.split(' ', 1)[0].split('@')
        if team_abbv == home_team:
            opp = away_team
            opp_excel = "vs. {}".format(away_team)
            home = True
        else:
            opp = home_team
            opp_excel = "at {}".format(home_team)
            home = False
        return opp, opp_excel, home


class QB(Player):
    """QB subclass of Player."""

    def __init__(self, player):
        self.assign(player)

        # pressure
        self.line_sack_rate = None
        self.opp_sack_rate = None

        # season stats
        self.rush_yds = None
        self.pass_dyar = None
        self.qbr = None

        # matchup stats
        self.opp_yds_att = None
        self.opp_comp_perc = None
        self.opp_td_perc = None

    def set_sack_fields(self, line_sack_rate, opp_sack_rate):
        self.line_sack_rate = line_sack_rate
        self.opp_sack_rate = opp_sack_rate

    def set_matchup_fields(self, opp_yds_att, opp_comp_perc, opp_td_perc):
        self.opp_yds_att = opp_yds_att
        self.opp_comp_perc = opp_comp_perc
        self.opp_td_perc = opp_td_perc

    def set_season_fields(self, rush_yds, pass_dyar, qbr):
        self.rush_yds = rush_yds
        self.pass_dyar = pass_dyar
        self.qbr = qbr
    # def set_matchup_fields(self, opp_)

    def get_writable_header(self):
        return ['Position', 'Name', 'Opp', 'Abbv',   # standard
                'Implied Total', 'O/U', 'Line',  # vegas
                'Rushing Yards', 'DYAR', 'QBR',  # season
                'O-Line Sack%', 'D-Line Sack%',  # pressure
                'Def Yds/Att', 'Def Comp%', 'Def TD%',  # matchup
                'Ave PPG', 'ECR', '+/- Rank',  # rankings
                'Salary', 'Salary%',  # draftkings
                'FD Salary', 'FD Salary%', 'FD +/- Rank',  # fdraft
                'ECR Data', 'Salary Rank', 'FDraft Salary Rank']  # hidden

    def get_writable_row(self):
        return [self.position, self.name, self.matchup, self.team_abbv,   # standard
                self.projected, self.overunder, self.line,  # vegas
                self.rush_yds, self.pass_dyar, self.qbr,  # season stats
                self.line_sack_rate, self.opp_sack_rate,  # pressure
                self.opp_yds_att, self.opp_comp_perc, self.opp_td_perc,  # matchup
                self.average_ppg, 'rank', '+/- r',  # rankings
                self.salary, self.salary_percent,  # draftkings
                self.fdraft_salary, self.fdraft_salary_perc, '+/- fd',  # fdraft
                self.rank, 'salaryrnk', 'fdraft rank']  # hidden

    def __repr__(self):
        return("QB({}, {} ({}), {})".format(self.name, self.salary, self.salary_percent, self.opp_excel))


class RB(Player):
    """RB subclass of Player."""

    def __init__(self, player):
        self.assign(player)

        # matchup
        self.run_dvoa = None
        self.rb_pass_dvoa = None
        self.oline_adj_line_yds = None
        self.opp_adj_line_yds = None

        # season
        self.season_snap_percent = None
        self.season_rush_atts = None
        self.season_targets = None

        # last week
        self.snap_percentage_by_week = None  # list
        self.rush_atts_weeks = None  # dict
        self.targets_weeks = None  # list
        # actual last week variables
        self.last_week_snap_percent = None
        self.last_week_rush_atts = None
        self.last_week_targets = None

    def set_dvoa_fields(self, run_dvoa, rb_pass_dvoa):
        self.run_dvoa = run_dvoa
        self.rb_pass_dvoa = rb_pass_dvoa

    def set_line_fields(self, oline_adj_line_yds, opp_adj_line_yds):
        self.oline_adj_line_yds = oline_adj_line_yds
        self.opp_adj_line_yds = opp_adj_line_yds

    def set_season_fields(self, season_snap_percent, season_rush_atts, season_targets):
        self.season_snap_percent = season_snap_percent
        self.season_rush_atts = season_rush_atts
        self.season_targets = season_targets

    def set_last_week_fields(self):
        self.last_week_snap_percent = self.get_last_week_snaps()
        self.last_week_rush_atts = self.get_last_week_rush_atts()
        self.last_week_targets = self.get_last_week_targets()

    def get_last_week_snaps(self):
        return self.snap_percentage_by_week[-1]

    def get_last_week_rush_atts(self):
        # print("from class: {}".format(self.rush_atts_weeks))
        return list(self.rush_atts_weeks.values())[-1]

    def get_last_week_targets(self):
        return self.targets_weeks[-1]

    def get_writable_header(self):
        return ['Position', 'Name', 'Opp', 'Abbv',   # standard
                'Implied Total', 'O/U', 'Line',  # vegas
                'Run DVOA', 'Pass DVOA', 'O-Line', 'D-Line',  # matchup
                'Snap%', 'Rush ATTs', 'Targets',  # season
                'Snap%', 'Rush ATTs', 'Targets',  # last week
                'Ave PPG', 'ECR', '+/- Rank',  # rankings
                'Salary', 'Salary%',  # draftkings
                'FD Salary', 'FD Salary%', 'FD +/- Rank',  # fdraft
                'ECR Data', 'Salary Rank', 'FDraft Salary Rank']  # hidden

    def get_writable_row(self):
        return [self.position, self.name, self.matchup, self.team_abbv,
                self.projected, self.overunder, self.line,  # vegas
                self.run_dvoa, self.rb_pass_dvoa, self.oline_adj_line_yds, self.opp_adj_line_yds,  # matchup
                self.season_snap_percent, self.season_rush_atts, self.season_targets,  # season
                self.last_week_snap_percent, self.last_week_rush_atts, self.last_week_targets,  # last week
                self.average_ppg, 'rank', '+/- r',   # rankings
                self.salary, self.salary_percent,  # draftkings
                self.fdraft_salary, self.fdraft_salary_perc, '+/- fd',  # fdraft
                self.rank, 'salaryrnk', 'fdraft rank']  # hidden

    def __repr__(self):
        return("RB({}, {} ({}), {})".format(self.name, self.salary, self.salary_percent, self.opp_excel))


class WR(Player):
    """WR subclass of Player."""

    def __init__(self, player):
        self.assign(player)

        # matchup
        self.pass_def_rank = None
        self.wr1_rank = None
        self.wr2_rank = None
        self.dline = None

        # season
        self.season_snap_percent = None
        self.season_targets = None
        self.season_recepts = None

        # last week
        self.snap_percentage_by_week = None  # list
        self.recepts_weeks = None  # dict
        self.targets_weeks = None  # list
        # actual last week variables
        self.last_week_snap_percent = None
        self.last_week_recepts = None
        self.last_week_targets = None

    def set_dvoa_fields(self, pass_def_rank, wr1_rank, wr2_rank):
        self.pass_def_rank = pass_def_rank
        self.wr1_rank = wr1_rank
        self.wr2_rank = wr2_rank

    def set_season_fields(self, season_snap_percent, season_targets, season_recepts):
        self.season_snap_percent = season_snap_percent
        self.season_targets = season_targets
        self.season_recepts = season_recepts

    def set_last_week_fields(self):
        self.last_week_snap_percent = self.get_last_week_snaps()
        self.last_week_recepts = self.get_last_week_recepts()
        self.last_week_targets = self.get_last_week_targets()

    def get_last_week_snaps(self):
        return self.snap_percentage_by_week[-1]

    def get_last_week_recepts(self):
        # print("from class: {}".format(self.rush_atts_weeks))
        return list(self.recepts_weeks.values())[-1]

    def get_last_week_targets(self):
        return self.targets_weeks[-1]

    def get_writable_header(self):
        return ['Position', 'Name', 'Opp', 'Abbv',   # standard
                'Implied Total', 'O/U', 'Line',  # vegas
                'Pass DVOA', 'vs. WR1', 'vs. WR2',  # matchup
                'Snap%', 'Targets', 'Recepts',  # season
                'Snap%', 'Targets', 'Recepts',  # last week
                'Ave PPG', 'ECR', '+/- Rank',  # rankings
                'Salary', 'Salary%',  # draftkings
                'FD Salary', 'FD Salary%', 'FD +/- Rank',  # fdraft
                'ECR Data', 'Salary Rank', 'FDraft Salary Rank']  # hidden

    def get_writable_row(self):
        return [self.position, self.name, self.matchup, self.team_abbv,  # standard
                self.projected, self.overunder, self.line,  # vegas
                self.pass_def_rank, self.wr1_rank, self.wr2_rank,  # matchup
                self.season_snap_percent, self.season_targets, self.season_recepts,  # season
                self.last_week_snap_percent, self.last_week_targets, self.last_week_recepts,  # last week
                self.average_ppg, 'rank', '+/- r',   # rankings
                self.salary, self.salary_percent,  # draftkings
                self.fdraft_salary, self.fdraft_salary_perc, '+/- fd',  # fdraft
                self.rank, 'salaryrnk', 'fdraft rank']  # hidden

    def __repr__(self):
        return("WR({}, {} ({}), {})".format(self.name, self.salary, self.salary_percent, self.opp_excel))


class TE(Player):
    """TE subclass of Player."""

    def __init__(self, player):
        self.assign(player)

        # matchup
        self.pass_def_rank = None
        self.te_rank = None

        # season
        self.season_snap_percent = None
        self.season_targets = None
        self.season_recepts = None

        # last week
        self.snap_percentage_by_week = None  # list
        self.recepts_weeks = None  # dict
        self.targets_weeks = None  # list
        # actual last week variables
        self.last_week_snap_percent = None
        self.last_week_recepts = None
        self.last_week_targets = None

    def set_dvoa_fields(self, pass_def_rank, te_rank):
        self.pass_def_rank = pass_def_rank
        self.te_rank = te_rank

    def set_season_fields(self, season_snap_percent, season_targets, season_recepts):
        self.season_snap_percent = season_snap_percent
        self.season_targets = season_targets
        self.season_recepts = season_recepts

    def set_last_week_fields(self):
        self.last_week_snap_percent = self.get_last_week_snaps()
        # print("set last_week_snap_percent to {}".format(self.last_week_snap_percent))
        self.last_week_recepts = self.get_last_week_recepts()
        # print("set last_week_snap_percent to {}".format(self.last_week_recepts))
        self.last_week_targets = self.get_last_week_targets()
        # print("set last_week_snap_percent to {}".format(self.last_week_targets))

    def get_last_week_snaps(self):
        return self.snap_percentage_by_week[-1]

    def get_last_week_recepts(self):
        # print("from class: {}".format(self.rush_atts_weeks))
        return list(self.recepts_weeks.values())[-1]

    def get_last_week_targets(self):
        return self.targets_weeks[-1]

    def get_writable_header(self):
        return ['Position', 'Name', 'Opp', 'Abbv',   # standard
                'Implied Total', 'O/U', 'Line',  # vegas
                'Pass DVOA', 'vs. TE',  # matchup
                'Snap%', 'Targets', 'Recepts',  # season
                'Snap%', 'Targets', 'Recepts',  # last week
                'Ave PPG', 'ECR', '+/- Rank',   # rankings
                'Salary', 'Salary%',  # draftkings
                'FD Salary', 'FD Salary%', 'FD +/- Rank',  # fdraft
                'ECR Data', 'Salary Rank', 'FDraft Salary Rank']  # hidden

    def get_writable_row(self):
        return [self.position, self.name, self.matchup, self.team_abbv,  # standard
                self.projected, self.overunder, self.line,  # vegas
                self.pass_def_rank, self.te_rank,  # matchup
                self.season_snap_percent, self.season_targets, self.season_recepts,  # season
                self.last_week_snap_percent, self.last_week_targets, self.last_week_recepts,  # last week
                self.average_ppg, 'rank', '+/- r',  # rankings
                self.salary, self.salary_percent,  # draftkings
                self.fdraft_salary, self.fdraft_salary_perc, '+/- fd',  # fdraft
                self.rank, 'salaryrnk', 'fdraft rank']  # hidden

    def __repr__(self):
        return("TE({}, {} ({}), {})".format(self.name, self.salary, self.salary_percent, self.opp_excel))


class DST(Player):
    """DST subclass of Player."""

    def __init__(self, player):
        self.assign(player)

    def __repr__(self):
        return("DST({}, {} ({}), {})".format(self.name, self.salary, self.salary_percent, self.opp_excel))

    def get_writable_header(self):
        return ['Position', 'Name', 'Opp', 'Abbv',  # standard
                'Implied Total', 'O/U', 'Line',  # vegas
                'Ave PPG', 'ECR', '+/- Rank',  # rankings
                'Salary', 'Salary%',  # draftkings
                'FD Salary', 'FD Salary%', 'FD +/- Rank',  # fdraft
                'ECR Data', 'Salary Rank', 'FDraft Salary Rank']  # hidden

    def get_writable_row(self):
        return [self.position, self.name, self.matchup, self.team_abbv,  # standard
                self.projected, self.overunder, self.line,  # vegas
                self.average_ppg, 'rank', '+/- r',  # rankings
                self.salary, self.salary_percent,  # draftkings
                self.fdraft_salary, self.fdraft_salary_perc, '+/- fd',  # fdraft
                self.rank, 'salaryrnk', 'fdraft rank']  # hidden
