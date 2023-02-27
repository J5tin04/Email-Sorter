def sort_email(cont_or_send, content, folder):
    target_folder = inbox.Folders(folder)

    for i in "12345":
        for message in inbox.Items:
            try:
                if cont_or_send == "1":
                    if message.SenderEmailAddress.startswith(content) or message.SenderEmailAddress.endswith(content):
                        message.Move(target_folder)
                else:
                    if (content in message.body) or (content in message):
                        message.Move(target_folder)
            except:
                continue

    print('done')


import win32com.client as client

outlook = client.Dispatch('Outlook.Application');

namespace = outlook.GetNameSpace('MAPI');

while True:
    email = input("Please enter your email: ");

    account = namespace.Folders[email];

    inbox = account.Folders['Inbox'];

    while True:
        print('\nPlease select an option:\n   1) Create New Folder\n   2) Sort Email\n   3) Exit')
        main_menu = input("   >> ")

        match main_menu:

            case "1":
                Folder_name = input("\n      Please input the folder's name (Press Q to quit/seperate multiple files using '|')>> ")

                if (Folder_name == "Q"):
                    continue

                else:
                    Folder_name = Folder_name.split('|');

                    for name in Folder_name:
                        try:
                            inbox.Folders.Add(name)
                            print("\n      ", name, " is created")
                        except:
                            print("\n      Error: Folder has been created befor")
                            continue

            case "2":
                while True:
                    print("\n      Please select one of the following:\n            1) Sort by sender\n            2) Sort by content\n            3) Back")
                    sort_menu = input("            >> ")

                    match sort_menu:
                        case "1":
                            sender = input("Plese enter email of sender ('q' to cancel): ")
                            foldername = input("Please enter name of folder ('q' to cancel): ")

                            if sender != 'q' or foldername != 'q':
                                sort_email('1', sender, foldername)

                        case "2":
                            content = input("Plese enter part of email content ('q' to cancel): ")
                            foldername = input("Please enter name of folder ('q' to cancel): ")

                            if content != 'q' or foldername != 'q':
                                sort_email('2', content, foldername)

                        case "3":
                            break
                        
                        case _:
                            print("      Please enter a valid input")


            case "3":
                while True:
                    confirmation = input("\n   Are you sure you want to exit (y/n): ")

                    if confirmation in ["Yes","yeS","YeS","YES","Y","yes","y"]:
                            exit()

                    elif confirmation in ["NO","No","nO","no",'n']:
                            continue
                        
                    else:
                        print("   Please input a valid input")

            case _:
                print("Please enter a valid input")

    exit()