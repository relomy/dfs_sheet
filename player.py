"""Object to hold various stats per player."""


class Player:
    """Creates Player object."""
    def __init__(self, position, player_name, team_abbv, salary, game_info, average_ppg):
        self.name = player_name

        # from DK salary CSV
        self.name = player_name
        self.position = position
        self.team_abbv = team_abbv
        self.salary = salary
        self.game_info = game_info
        self.average_ppg = average_ppg

        self.opponent, self.opp_excel, self.home_team = self.get_opponent_matchup(game_info, team_abbv)

        # calculate salary percent
        self.salary_percent = "{0:.1%}".format(float(salary) / 50000)

        # ECR
        self.matchup = None

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

    def __repr__(self):
        return "Player({}, {})".format(self.name, self.rank)

    def set_ecr_fields(self, matchup, rank):
        self.matchup = matchup
        self.rank = rank

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
    def __init__(self, position, player_name, team_abbv, salary, game_info, average_ppg):
        Player.__init__(self, position, player_name, team_abbv, salary, game_info, average_ppg)

    def __repr__(self):
        return("QB({}, {} ({}), {})".format(self.name, self.salary, self.salary_percent, self.opp_excel))


class RB(Player):
    """RB subclass of Player."""
    def __init__(self, position, player_name, team_abbv, salary, game_info, average_ppg):
        Player.__init__(self, position, player_name, team_abbv, salary, game_info, average_ppg)

        # matchup
        self.run_dvoa = None
        self.pass_dvoa = None
        self.oline = None
        self.dline = None

        # season
        self.season_snap_percent = None
        self.season_rush_atts = None
        self.season_targets = None

        # last week
        self.week_snap_percent = None
        self.week_rush_atts = None
        self.week_targets = None

    def set_dvoa_fields(self, run_dvoa, rb_pass_dvoa):
        self.run_dvoa = run_dvoa
        self.rb_pass_dvoa = rb_pass_dvoa

    def __repr__(self):
        return("RB({}, {} ({}), {})".format(self.name, self.salary, self.salary_percent, self.opp_excel))


class WR(Player):
    """WR subclass of Player."""
    def __init__(self, position, player_name, team_abbv, salary, game_info, average_ppg):
        Player.__init__(self, position, player_name, team_abbv, salary, game_info, average_ppg)

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
        self.week_snap_percent = None
        self.week_targets = None
        self.week_recepts = None

    def set_dvoa_fields(self, pass_def_rank, wr1_rank, wr2_rank):
        self.pass_def_rank = pass_def_rank
        self.wr1_rank = wr1_rank
        self.wr2_rank = wr2_rank

    def __repr__(self):
        return("wr({}, {} ({}), {})".format(self.name, self.salary, self.salary_percent, self.opp_excel))


class TE(Player):
    """TE subclass of Player."""
    def __init__(self, position, player_name, team_abbv, salary, game_info, average_ppg):
        Player.__init__(self, position, player_name, team_abbv, salary, game_info, average_ppg)

    def __repr__(self):
        return("TE({}, {} ({}), {})".format(self.name, self.salary, self.salary_percent, self.opp_excel))


class DST(Player):
    """DST subclass of Player."""
    def __init__(self, position, player_name, team_abbv, salary, game_info, average_ppg):
        Player.__init__(self, position, player_name, team_abbv, salary, game_info, average_ppg)

    def __repr__(self):
        return("DST({}, {} ({}), {})".format(self.name, self.salary, self.salary_percent, self.opp_excel))
