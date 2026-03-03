class User:
    def __init__(self, name, email, year, number,discord,academicFocus, major, preferences, answers):

     self.name = name
     self.email = email
     self.year = year
     self.number = number
     self.discord = discord
     self.academicFocus = academicFocus
     self.major = major
     self.preferences = preferences #False: No preference for Major  True: Same major
     self.answers = answers
     self.score = 0


    @property
    def getName(self):
       return self.name
    
    @property
    def getEmail(self):
       return self.email

    @property
    def getYear(self):
       return self.year

    @property
    def getNumber(self):
       return self.number

    @property
    def getDiscord(self):
       return self.discord

    @property
    def getMajor(self):
       return self.major

    @property
    def getPreferences(self):
       return self.preferences
    
    @property
    def getAnswers(self):
       return self.answers

    @property
    def getScores(self):
       return self.score
    
    @property
    def getAcademicFocus(self):
       return self.academicFocus

    def setScore(self, scores):
       self.score = scores

    
    
