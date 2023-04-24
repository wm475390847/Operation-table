from packs.attr_display import AttrDisplay


class Rule(AttrDisplay):
    def __init__(
        self, en, formula_mode, three_NC, diamond_1, diamond_2, diamond_3, diamond_4
    ):
        self.en = en
        self.formula_mode = formula_mode
        self.three_NC = three_NC
        self.diamond_1 = diamond_1
        self.diamond_2 = diamond_2
        self.diamond_3 = diamond_3
        self.diamond_4 = diamond_4
