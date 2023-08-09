import itertools
from string import digits, punctuation, ascii_letters


# symbols = digits + punctuation + ascii_letters
# print(symbols)


def brute_excel_doc():
    print("Hello friend!!!")
    try:
        password_length = input("write the length of the password from X to Y, for example 3 - 7: ")
        password_length =  [int(item) for item in password_length.split("-")] 
        
    except:
        print("Check the information you wrote pls!")

    print("If password includes only digits type - 1\nIf password includes only letters - 2\n"
          "If only digits and letters - 3\nIf digits , punctuation and ascii_letters - 4")
    
    try:
        choice = int(input(": "))
        if choice == 1:
            possible_symbols = digits
        elif choice == 2:
            possible_symbols = ascii_letters
        elif choice == 3:
            possible_symbols = digits + ascii_letters
        elif choice == 4:
            possible_symbols = digits + ascii_letters + punctuation
        else:
            print("What do you want uncle Sam?")
        print(possible_symbols)
    except:
        print("What do you want uncle Sam?")
        
    # brute excel doc
    for pass_length in range(password_length[0], password_length[1] + 1):
        for password in itertools.product(possible_symbols, repeat=pass_length):
            password = "".join(password)
            print(password)

def main():
    brute_excel_doc()


if __name__ == "__main__":
    main()
