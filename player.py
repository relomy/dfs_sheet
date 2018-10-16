"""Object to hold various stats per player."""


class Player:
    """Creates Player object."""

    def __init__(self, player_name):
        self.name = player_name

    def __repr__(self):
        return "Player({})".format(self.name)
