from user import User

class Mentor(User):
    count = 0
    def __init__(self, name, email, year, number,discord,academicFocus, major, preferences,answers, numMentees):
        super().__init__(name,email, year, number,
        discord, academicFocus, major, preferences, answers)
        self.numMentees = numMentees
        Mentor.count += 1

    @property
    def getNumMentees(self):
        return self.numMentees
