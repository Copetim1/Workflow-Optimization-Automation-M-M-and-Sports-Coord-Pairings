from user import User

class Mentee(User):
    count = 0
    def __init__(self, name, email, year, number,discord, major, sameMajor, preferences, answers):
        super().__init__(name,email, year, number,
        discord, major, sameMajor, preferences, answers)
        count += 1

      