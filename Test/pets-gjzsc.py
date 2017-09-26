def describe_pet(animal_type,pet_name):
    """显示宠物信息"""
    print("\nI have a "+animal_type+".")
    print('My '+animal_type+"'s name is "+pet_name.title()+".")

describe_pet(animal_type='hamster',pet_name='harry')
describe_pet(pet_name='willie',animal_type='dog')