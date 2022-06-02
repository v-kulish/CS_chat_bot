import win32com.client
import re
import pathlib
from send import SendMail
import pandas as pd
import os
import openpyxl


class CsBot:
    negative_responses = ('quit', 'close', 'pause', 'off', 'exit',
                          'nothing', 'bye', 'no', 'no, thanks', 'no, thank you')
    matching_phrases = {'luggage_storage': [r'.*luggage.*(basel)', r'.*store.*luggage', r'.*keep.*luggage'],
                        'parking_intent': [r'.*(no).*park', r'.*parking.*(no)t', r'.*(park).*car', r'.*(parking)'],
                        'send_ci_info': [r'.*send.*check.*info', r'.*send.*ci.*details', r'.*send.*ci.*details', r'.*send.*ci.*info'],
                        'early_ci_general': [r'.*early.*ci.*general', r'.*early.*ci.*options', r'.*early.*check.*mail'],
                        'early_ci_tomorrow': [r'.*early.*ci.*(no)t', r'.*(no).*early.*ci', r'.*early.*ci.*tomorrow', r'.*early.*ci.*today', r'.*early.*ci.*now'],
                        'late_co_general': [r'.*late.*co.*general', r'.*late.*co.*options', r'.*late.*check.*mail'],
                        'late_co_tomorrow': [r'.*late.*co.*(no)t', r'.*(no).*late.*co.*today', r'.*late.*co.*now'],
                        'early_ci_and_late_co': [r'.*late.*co.*and.*early.*ci', r'.*early.*ci.*and.*late.*co'],
                        'general_access_info': [r'.*how.*to.*get.*to.*apart', r'.*how.*access', r'.*when.*code', r'.*when.*receive'],
                        'arriving_late': [r'.*late.*arrival', r'.*arrive.*late'],
                        'prol_not_possible':[r'.*prol.*not', r'.*no.*prol'],
                        'id_but_no_form': [r'.*id.*but.*no.*form', r'.*id.*no.*form', '.*id.*but.*form.*no'],
                        'tech_and_inventory': [r'.*tech.*inventory', r'.*inventory.*tech'],
                        'technician_task': [r'.*technician'],
                        'inventory_check': [r'.*inventory']

                        }
    def __init__(self):
        self.welcome()
    def welcome(self):
        self.name = input("I will be your assistant with the questions from customers. Could you tell me your name? ")
        print('''A short guide:\n 
         The bot can provide a bunch of templates:
         - early CI or late CO (in general or now - available or not);
         - parking (available or not);
         - luggage (Basel or other locations);
         - is it okay to arrive late (yes);
         - when prolongation is not possible;
         - the guest sent an id, but without city tax form;
         - responding to an inventory check;
         - confirming that technicians have been informed;
         - confirming both inventory and task for technicians.\n
         It can also send multiple CI templates for Booking.com and Expedia reservations.\n''')

        will_help = input(f"So, {self.name}, what can I help you with? ")

        if will_help.lower() in self.negative_responses:
            print("Ok, start the program again if you need help!")
            return

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
                        return input('Looks like something went wrong. Please try again or use a different command. ')
                if found_match and key == 'early_ci_general':
                    return self.early_ci_general()
                if found_match and key == 'early_ci_tomorrow':
                    try:
                        if found_match.groups()[0] == 'no':
                            return self.early_ci_now(False)
                    except:
                        return self.early_ci_now()
                if found_match and key == 'late_co_general':
                    return self.late_co_general()
                if found_match and key == 'late_co_tomorrow':
                    try:
                        if found_match.groups()[0] == 'no':
                            return self.late_co_now(False)
                    except:
                        return self.late_co_now()
                if found_match and key == 'early_ci_and_late_co':
                    return self.early_ci_and_late_co()
                if found_match and key == 'general_access_info':
                    return self.general_access_info()
                if found_match and key == 'arriving_late':
                    return self.arriving_late()
                if found_match and key == 'prol_not_possible':
                    return self.prol_not_possible()
                if found_match and key == 'id_but_no_form':
                    return self.id_but_no_form()
                if found_match and key == 'inventory_check':
                    return self.inventory_check()
                if found_match and key == 'tech_and_inventory':
                    return self.tech_and_inventory()
                if found_match and key == 'technician_task':
                    return self.technician_task()
                if found_match and key == 'second task':
                    pass
                if found_match and key == 'second task':
                    pass
        return input(f'I didn\'t understand you, {self.name}. Could you rephrase your command? ')

    def luggage_storage_intent(self, city=''):

        if city == 'basel':
            f = open(r'templates\luggage_basel.txt', 'r')
            text = f.read()
            print(text)
            f.close()
            return input(f'Can I help you with any other question, {self.name}? ')
        f = open(r'templates\no_luggage.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def parking_intent(self, available=True):
        if available:
            f = open(r'templates\parking.txt', 'r')
            text = f.read()
            print(text)
            f.close()
            return input(f'Can I help you with any other question, {self.name}? ')
        f = open(r'templates\no_parking.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def send_ci_info_intent(self, mail_number=''):
        if mail_number == '':
            send_mails = SendMail()
        else:
            send_mails = SendMail(int(mail_number))
        return input(f'Can I help you with any other question, {self.name}? ')

    def early_ci_general(self):
        f = open(r'templates\early_ci_general.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def early_ci_now(self, possible=True):
        if possible:
            f = open(r'templates\early_ci_possible.txt', 'r')
            text = f.read()
            print(text)
            f.close()
            return input(f'Can I help you with any other question, {self.name}? ')
        f = open(r'templates\early_ci_not.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def late_co_general(self):
        f = open(r'templates\late_co_general.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def late_co_now(self, possible=True):
        if possible:
            f = open(r'templates\late_co_possible.txt', 'r')
            text = f.read()
            print(text)
            f.close()
            return input(f'Can I help you with any other question, {self.name}? ')
        f = open(r'templates\late_co_not.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def early_ci_and_late_co(self):
        f = open(r'templates\early_ci_and_late_co.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def general_access_info(self):
        f = open(r'templates\how_to_access.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def arriving_late(self):
        f = open(r'templates\late_arrival.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def prol_not_possible(self):
        f = open(r'templates\prol_not_possible.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def id_but_no_form(self):
        f = open(r'templates\id_but_no_form.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def inventory_check(self):
        f = open(r'templates\inventory_check.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def tech_and_inventory(self):
        f = open(r'templates\inventory_and_tech.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')

    def technician_task(self):
        f = open(r'templates\inform_tech.txt', 'r')
        text = f.read()
        print(text)
        f.close()
        return input(f'Can I help you with any other question, {self.name}? ')


chat = CsBot()
