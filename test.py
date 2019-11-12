print("Would you like to print the Duplicates List as an Excel File?\n\nPlease type Y or N:")
UserInput = str(input())

print(UserInput)

user_input_steps = 0

while user_input_steps == 0:
    if UserInput == "Y":
        print("Making Excel File")
        user_input_steps += 1
    if UserInput == "N":
        print("Not creating Excel File")
        user_input_steps += 1
    else:
        print("Not valid input")