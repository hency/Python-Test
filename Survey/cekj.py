import sys
class CC:
    def __init__(self):
        self.xx=1
        self.setXY(6,5)
        # print(self.xx)
        # print(self.x)
    def setXY(self, x, y):
        self.x = x
        self.y = y
        self.xx=4

    # def printXY(self):
    #     print(self.x, self.y)


dd = CC()
# print(dd.x)
current = sys.stdout
print(current)
# print(dd.xx)
# dd.setXY(4, 5)
# print(dd.xx)
# dd.printXY()