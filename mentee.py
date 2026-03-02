from user import User

class Mentee(User):
    count = 0
    def __init__(self, name, email, year, number,discord,academicFocus, major, preferences, answers):
        super().__init__(name,email, year, number,
        discord,academicFocus, major, preferences, answers)
        Mentee.count += 1

      