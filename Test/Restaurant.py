class Restaurant():

    def __init__(self,restaurant_name,cuisine_type):
        self.restaurant_name = restaurant_name
        self.cuisine_type = cuisine_type

    def describe_restaurant(self):
        long_name = self.restaurant_name + " " + self.cuisine_type
        return long_name.title()

    def open_restaurant(self):
        print(self.restaurant_name + "正在营业。")

restaurant_one = Restaurant("小柴米","中餐")
print(restaurant_one.describe_restaurant())
restaurant_one.open_restaurant()