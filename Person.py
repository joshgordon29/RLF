class Person:
    def __init__(self, first, last, location, capability, market, email, is_mentor):
        self.first = first.strip()
        self.last = last.strip()
        self.full_name = self.first + ' ' + self.last
        self.location = location.strip()
        self.capability = capability.strip()
        self.market = market.strip()
        self.email = email.strip()
        self.is_mentor = is_mentor

    def __eq__(self, other):
        return (
            self.location == other.location and
            self.capability == other.capability and
            self.market == other.market)
