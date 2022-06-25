import win32com.client
import re

class Cancellations:

    def __init__(self):

        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  # to trigger outlook application
        choose_folder = input("\nWhich folder should I scan? \n'2) booking.com/ expedia' or '4) CUSTOMER SERVICE'? \n(press 2 or 4)>>> ")
        while choose_folder != '2' and choose_folder != '4':
            choose_folder = input("Press '2' for '2) booking.com/ expedia' or press '4' for '4) CUSTOMER SERVICE'\n>>> ")
        self.folder_choice = {'2': '2) booking.com/ expedia', '4': '4) CUSTOMER SERVICE'}
        folder_to_check = self.folder_choice[choose_folder]
        self.get_current_list()
        self.check_folder(folder_to_check)
        qst_if_check = input("\nDo you want to check the other folder, too? (Y/N)\n>>> ").lower()
        if 'y' in qst_if_check:
            choose_folder = '2' if choose_folder == '4' else '4'
            folder_to_check = self.folder_choice[choose_folder]
            self.check_folder(folder_to_check)

    def get_current_list(self):

        folder = self.outlook.Folders.Item("Customer Service")
        unassigned = folder.Folders.Item("Inbox").Folders.Item("1) magarental nieprzypisane")

        bookings = unassigned.Items
        booking = bookings.GetLast()
        # message is treated as each mail in for loop

        self.res_list = {}
        count = 1
        thresh = 100
        print('\nCollecting all numbers of unassigned reservations first. This will take a few minutes.\n')
        for booking in bookings:
            if not re.match('NEW BOOKING', booking.Subject):
                continue
            try:
                b_num = re.findall("Magarental ID.*([0-9]{7})", str(booking.Body))[0]
            except:
                continue
            res_date = re.findall("From: (\d+ \w+)", str(booking.Subject))[0]
            self.res_list[b_num] = res_date
            count += 1
            if count == thresh:
                print(f'found {count} reservations')
                thresh +=100

    def check_folder(self, folder_to_check):

        folder = self.outlook.Folders.Item("Customer Service")
        inbox = folder.Folders.Item("Inbox").Folders.Item(folder_to_check)
        done = folder.Folders.Item("Inbox").Folders.Item("2) booking.com/ expedia")
        unassigned = folder.Folders.Item("Inbox").Folders.Item("1) magarental nieprzypisane")

        messages = inbox.Items
        bookings = unassigned.Items
        message = messages.GetLast()
        booking = bookings.GetLast()

        for message in messages:
            if not (re.match('CANCELLATION -  From', message.Subject) and not "AIRBNB" in str(message.Body)):
            # based on the subject replying to email
                continue

            try:
                b_num = re.findall("Magarental ID.*([0-9]{7})", str(message.Body))[0]
            except:
                continue
            if not b_num in self.res_list.keys():
                print(f"Reservation {b_num} not in unassigned")
                continue
            print(f'looking for {b_num}')
            filter_str = f"@SQL=urn:schemas:httpmail:subject like '%{self.res_list[b_num]}%'"
            sub_bookings = bookings.Restrict(filter_str)
            sub_booking = sub_bookings.GetLast()
            for sub_booking in sub_bookings:
                if not (re.match('NEW BOOKING', sub_booking.Subject) and b_num in str(sub_booking.Body)):
                    continue
                sub_booking.Move(done)
                print(f'Booking {b_num} moved from "1) magarental nieprzypisane" to "2) booking.com/ expedia"')
                quit_c = input('Quit? (Y/N)>>> ').lower()
                if 'y' in quit_c:
                    return None
