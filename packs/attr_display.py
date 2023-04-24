class AttrDisplay(object):
    def __str__(self):
        result = []
        for key in sorted(self.__dict__):
            result.append("{}={}".format(key, getattr(self, key)))
        return "{}({})".format(self.__class__.__name__, ", ".join(result))
