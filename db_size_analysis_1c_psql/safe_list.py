class SafeList(list):
    def get(self, index, default=None):
        try:
            return self[index]
        except IndexError:
            return default
