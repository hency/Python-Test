from random import randint

class Die():
    """随机掷骰子"""

    def __init__(self,sides = 6):
        self.sides = sides

    def roll_die(self):
        x = randint(1,self.sides)
        print(x)
        
die_one = Die(6)
die_one.roll_die()