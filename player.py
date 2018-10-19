"""Object to hold various stats per player."""


class Player:
    """Creates Player object."""

    def __init__(self, player_name, matchup, rank):
        self.name = player_name
        self.matchup = matchup
        self.rank = rank

    def __repr__(self):
        return "Player({}, {})".format(self.name, self.rank)
