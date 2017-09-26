class Car():
    """一次模拟汽车的简单尝试"""

    def __init__(self,make,model,year):
        """初始化描述汽车的属性"""
        self.make = make
        self.model = model
        self.year = year

    def get_descriptive_name(self):
        """返回整洁的描述信息"""
        long_name = str(self.year)+' '+self.make+' '+self.model
        return long_name.title()

class Battery():
    """"一次模拟电动汽车电池的简单尝试"""

    def __init__(self,battery_size = 70):
        """初始化电瓶的属性"""
        self.battery_size = battery_size
    
    def describe_battery(self):
        """打印一条描述电瓶容量的消息"""
        print("这辆车拥有"+str(self.battery_size)+"-KWh的电池。")

class ElectricCar(Car):
    """电动汽车的独特之处"""
    def __init__(self,make,model,year):
        """初始化父类属性，再初始化电动汽车特有的属性"""
        super().__init__(make,model,year)
        self.battery = Battery()

my_tesla = ElectricCar('tesla','model s',2016)
print(my_tesla.get_descriptive_name())
my_tesla.battery.describe_battery()   
my_new_car = Car('audi','a4',2016)

print(my_new_car.get_descriptive_name())