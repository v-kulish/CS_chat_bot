import win32com.client
import re
import pathlib
from send import SendMail
from cancel import Cancellations
import pandas as pd
import os
import openpyxl


class CsBot:

    #quit phrases, must be matched exactly
    negative_responses = ('quit', 'close', 'pause', 'off', 'exit',
                          'nothing', 'bye', 'no', 'no, thanks', 'no, thank you')
    #dictionary for intent matching
    matching_phrases = {'luggage_storage': [r'.*luggage.*(basel)', r'.*luggage'],
                        'where_parking': [r'.*(location|parking).*(location|parking)', r'.*(list|parking).*(list|parking)'],
                        'parking_intent': [r'.*(no).*park', r'.*parking.*(no)t', r'.*(park).*car', r'.*(parking)'],
                        'send_ci_info': [r'.*send.*check.*info', r'.*ci.*info', r'.*ci.*details', r'.*send.*ci'],
                        'early_ci_general': [r'.*(general|options|template).*early.*ci', r'.*early.*ci.*(general|options|template)', r'.*early.*check.*mail'],
                        'early_ci_tomorrow': [r'.*early.*ci.*(no)t', r'.*(no).*early.*ci', r'.*early.*ci.*tomorrow', r'.*early.*ci.*today', r'.*early.*ci.*possib', r'.*early.*ci.*now', r'.*early.*ci.*ok', r'.*early.*ci.*1.*pm'],
                        'late_co_general': [r'.*late.*co.*general', r'.*late.*co.*options', r'.*late.*check.*mail', r'late.*co.*template'],
                        'late_co_tomorrow': [r'.*late.*co.*(no)t', r'.*(no).*late.*co.*today', r'.*late.*co.*now', r'.*late.*co.*ok', r'.*late.*co.*12', r'.*late.*co.*possib'],
                        'early_ci_and_late_co': [r'.*late.*co.*and.*early.*ci', r'.*early.*ci.*and.*late.*co'],
                        'general_access_info': [r'.*how.*to.*get.*to.*apart', r'.*how.*access', r'.*when.*code', r'.*when.*receive'],
                        'arriving_late': [r'.*late.*arriv.*', r'.*arrive.*late'],
                        'prol_not_possible':[r'.*prol.*not', r'.*no.*prol'],
                        'id_but_no_form': [r'.*id.*but.*no.*form', r'.*id.*no.*form', '.*id.*but.*form.*no'],
                        'tech_list': [r'.*(assign|task|who).*(assign|task|who).*(assign|task|who)', r'.*(tech|list).*(tech|list)'],
                        'tech_and_inventory': [r'.*tech.*inventory', r'.*inventory.*tech'],
                        'technician_task': [r'.*technician', r'.*tech.*task'],
                        'inventory_check': [r'.*inventory', r'.*clean.*inform'],
                        'id_and_payment_reminder': [r'.*id.*pay.*remind.*', r'.*remind.*id.*pay', r'.*need.*pay.*id', r'.*id.*pay.*need', r'.*pay.*id.*need'],
                        'id_confirm': [r'.*id.*(confirm|thank|receive)', r'.*(confirm|thank|receive).*id'],
                        'id_reminder': [r'.*id.*remind.*', r'.*remind.*id', r'.*id.*need', r'.*need.*id'],
                        'payment_reminder': [r'.*pay.*remind.*', r'.*remind,*pay', r'.*need.*pay', r'.*pay.*need', r'.*(missing|payment).*(missing|payment)'],
                        'emergency_numbers': [r'.*emerg.*number'],
                        'extra_services': [r'.*(extra.*bed|baby.*cot).*price', r'.*price.*(extra.*bed|baby.*cot)', r'.*extra.*service'],
                        'cancellations': [r'.*(cancel|check).*(cancel|check)'],
                        'custom_temp1': [r'.*custom.*1'],
                        'custom_temp2': [r'.*custom.*2'],
                        'custom_temp3': [r'.*custom.*3']

                        }
    def __init__(self):
        self.welcome()
    def welcome(self):
    # the main body of the chat function
        self.name = input("I will be your assistant with questions from customers. Could you tell me your name?\n>>>")
        self.guide = '''\n~ A short guide:\n
         The bot can provide a bunch of templates:
         - early CI or late CO (in general or now - available or not);
         - both early CI and late CO general;
         - parking (available or not);
         - luggage (Basel or other locations);
         - is it okay to arrive late (yes);
         - when prolongation is not possible;
         - the guest sent an id, but without city tax form;
         - responding to an inventory check;
         - confirming that technicians have been informed;
         - confirming both inventory and task for technicians;
         - ID received ('thank you' template)
         - ID reminder template;
         - payment template;
         - ID and payment template;
         - 3 custom templates (edit .txt file from "templates" and call it with "custom number*" command).\n
         It can provide certain information:
         - parking availability at our locations;
         - OTA extra services price list (extra bed/baby cot);
         - list of technicians (to whom a task should be assigned);
         - list of emergency numbers by city.\n
         It can find cancelled reservations and remove Magarental confirmations for them from 'unassigned':
         -> ask the program for a "cancellation check";\n
         It can also send multiple CI templates for Booking.com and Expedia reservations.
         -> copy you MR confirmation e-mails from Outlook;
         -> paste them in "bookings" folder;
         -> tell the program that you want to send ci infos and follow the instructions.\n'''
        print(self.guide)

        will_help = input(f"So, {self.name}, what can I help you with?\n>>>")

        if will_help.lower() in self.negative_responses:
            print("Ok, start the program again if you need help!")
            return
        # calling the loop chat method:
        self.handle_conversation(will_help.lower())

    def handle_conversation(self, reply):
        while not self.make_exit(reply.lower()):
            reply = self.match_reply(reply)

    def make_exit(self, reply):
        if reply in self.negative_responses:
            print("Ok, start the program again if you need help!")
            return True
        return False

    def match_reply(self, reply):
    #main method, matches the intent with the action
        for key, values in self.matching_phrases.items():
            for regex_pattern in values:
                found_match = re.match(regex_pattern, reply.lower())
                if found_match and key == 'luggage_storage':
                    try:
                        if found_match.groups()[0] == 'basel':
                            return self.luggage_storage_intent(found_match.groups()[0])
                    except:
                        return self.luggage_storage_intent()
                if found_match and key == 'parking_intent':
                    if found_match.groups()[0] == 'no':
                        return self.parking_intent(False)
                    return self.parking_intent(True)
                if found_match and key == 'send_ci_info':
                    try:
                        return self.send_ci_info_intent()
                    except:
                        return input('''Looks like something went wrong (or you have stopped the script). Please try again or use a different command.\n>>>''')
                if found_match and key == 'early_ci_tomorrow':
                    try:
                        if found_match.groups()[0] == 'no':
                            return self.early_ci_now(False)
                    except:
                        return self.early_ci_now()
                if found_match and key == 'late_co_tomorrow':
                    try:
                        if found_match.groups()[0] == 'no':
                            return self.late_co_now(False)
                    except:
                        return self.late_co_now()
                if found_match and key == 'cancellations':
                    try:
                        return self.cancellations()
                    except:
                        return input('''Looks like you have stopped the script (or something went wrong). Can I help you with any other question?\n>>>''')
                elif found_match:
                    return self.load_template(key)


        return input(f'I didn\'t understand you, {self.name}. Could you rephrase your command?\n ')

    def luggage_storage_intent(self, city=''):

        if city == 'basel':
            f = open(r'templates\luggage_basel.txt', 'r')
            text = f.read()
            print(text)
            f.close()
            return input(f'Can I help you with any other question, {self.name}?\n>>>')
        f = open(r'templates\no_luggage.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}?\n>>>')

    def load_template(self, key):
        file_loc = r'templates\\' + f'{key}.txt'
        f = open(file_loc, 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}?\n>>>')

    def parking_intent(self, available=True):
        if available:
            f = open(r'templates\parking.txt', 'r')
            text = f.read()
            print(text)
            f.close()
            return input(f'Can I help you with any other question, {self.name}?\n>>>')
        f = open(r'templates\no_parking.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}?\n>>>')

    def send_ci_info_intent(self, mail_number=''):
        if mail_number == '':
            send_mails = SendMail()
        else:
            send_mails = SendMail(int(mail_number))
        return input(f'Can I help you with any other question, {self.name}?\n>>>')

    def cancellations(self):
        cancellations = Cancellations()
        return input(f'Can I help you with any other question, {self.name}?\n>>>')


    def early_ci_now(self, possible=True):
        if possible:
            f = open(r'templates\early_ci_possible.txt', 'r')
            text = f.read()
            print(text)
            f.close()
            return input(f'Can I help you with any other question, {self.name}?\n>>>')
        f = open(r'templates\early_ci_not.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}?\n>>>')


    def late_co_now(self, possible=True):
        if possible:
            f = open(r'templates\late_co_possible.txt', 'r')
            text = f.read()
            print(text)
            f.close()
            return input(f'Can I help you with any other question, {self.name}?\n>>>')
        f = open(r'templates\late_co_not.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}?\n>>>')


chat = CsBot()
