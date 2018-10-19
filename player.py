"""Object to hold various stats per player."""


class Player:
    """Creates Player object."""
    def __init__(self, position, player_name, abbv, salary, game_info, average_ppg):
        self.name = player_name

        # from DK salary CSV
        self.name = player_name
        self.position = position
        self.abbv = abbv
        self.salary = salary
        self.game_info = game_info
        self.average_ppg = average_ppg

        # ECR
        self.matchup = None
        self.rank = None

        # vegas
        self.overunder = None
        self.line = None
        self.projected = None

    def __repr__(self):
        return "Player({}, {})".format(self.name, self.rank)

    def set_ecr_fields(self, matchup, rank):
        self.matchup = matchup
        self.rank = rank

    def set_vegas_fields(self, overunder, line, projected):
        self.overunder = overunder
        self.line = line
        self.projected = projected


class QB(Player):
    """Position specific subclass of Player."""
    def __init__(self, position, player_name, abbv, salary, game_info, average_ppg):
        Player.__init__(self, position, player_name, abbv, salary, game_info, average_ppg)

    def __repr__(self):
        return("RB({}, {}. {})".format(self.name, self.salary, self.game_info))
