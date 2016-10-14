class Time:
    def __init__(self, hours=0, minutes=0, sign=1):

        hours, minutes = self._organize(hours, minutes)

        self.hours = hours
        self.minutes = minutes
        self.sign = sign

    @classmethod
    def from_epochs(cls, epochs=0, sign=1):
        obj = cls(0, 0, sign=sign)
        obj.epochs = epochs
        return obj

    @classmethod
    def from_deftime(cls, time):
        mn = time.minute
        hr = time.hour
        return cls(hours=hr, minutes=mn, sign=1)

    @staticmethod
    def _organize(hours=0, minutes=0):
        if minutes >= 60:
            hours += minutes // 60
            minutes %= 60
        return hours, minutes

    @property
    def epochs(self):
        return (self.hours * 60 + self.minutes) * self.sign

    @epochs.setter
    def epochs(self, x):
        if x < 0:
            self.sign = -1
            x = -x
        self.hours = x // 60
        x %= 60
        self.minutes = x

    def __add__(self, other):
        epochs = self.epochs * self.sign + other.epochs * other.sign
        return Time.from_epochs(epochs)

    def __sub__(self, other):
        epochs = self.epochs * self.sign - other.epochs * other.sign
        if epochs < 0:
            epochs = -epochs
            return Time.from_epochs(epochs, sign=-1)
        else:
            return Time.from_epochs(epochs)

    def __lt__(self, other):
        return self.epochs < other.epochs

    def __le__(self, other):
        return self.epochs <= other.epochs

    def __gt__(self, other):
        return self.epochs > other.epochs

    def __ge__(self, other):
        return self.epochs >= other.epochs

    def __eq__(self, other):
        if other == "n.a.":
            return False
        return self.epochs == other.epochs

    def __ne__(self, other):
        if other == "n.a.":
            return True
        return self.epochs != other.epochs

    def __repr__(self):
        hr = str(self.hours)
        if len(hr) == 1:
            hr = "0" + hr
        mn = str(self.minutes)
        if len(mn) == 1:
            mn = "0" + mn
        return "{}:{}".format(hr, mn)

    def __str__(self):
        return self.__repr__()