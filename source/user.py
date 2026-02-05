class User:
    def __init__(self,name, email, year, number,discord, major, preferences):

     self.name = name
     self.email = email
     self.year = year
     self.number = number
     self.discord = discord
     self.major = major
     self.preferences = preferences
     self.answers = []

    def extractAnswers(self, answers):
        for i in (answers):
            self.answers.append(i)