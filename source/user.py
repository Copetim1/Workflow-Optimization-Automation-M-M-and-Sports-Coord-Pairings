class User:
    def __init__(self,name, email, year, number,discord, major, sameMajor, preferences, answers):

     self.name = name
     self.email = email
     self.year = year
     self.number = number
     self.discord = discord
     self.major = major
     self.sameMajor = sameMajor
     self.preferences = preferences
     self.answers = []

    def extractAnswers(self, answers):
        x = 5