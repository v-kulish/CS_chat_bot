import win32com.client
import re
import pathlib
import pandas as pd
import os
import openpyxl


class SendMail:
    client_address = ''
    bookings_names = []
    my_address = ''
    expedia = False
    city = ''
    card = ''
    lucerne_card = '''If you would like to receive the Visitor Card Lucerne,
please write us an email to customer.service@visionapartments.com
and we will prepare the card for you! <br><br>'''
    vevey_card = '''
    <b>Montreux Riviera Card</b><br>
    You have the option to receive a free Montreux Riviera Card! This
    provides you free access to transportation and discounts for many museums
    in the area. As each card has to be personalized, you have to request it
    directly at our local office from 9AM – 6PM Mo-Fr. If you are arriving
    during the weekend, you will have to inform the local office team in
    advance, as the office is closed during these ours. You can also send us
    an email and we will prepare it accordingly. We can only give the Riviera
    Card to our guests who pay city tax. If you only pay city tax for one
    person, only one person can be registered for the card.<br><br>
    '''
    requested = """<b>Please note that your arrival is only guaranteed if prior
     to check-in, we receive for each guest:<br><br>"""
    no_id = """&emsp;•    Copy of the ID (*BOTH sides) or Passport:</b>
    please send it to customer.service@visionapartments.com<br><br>"""
    no_p = "&emsp;•    <b>Payment:</b> you will receive a separate e-mail with you invoice<br><br>"
    no_f = """&emsp;•    <b>A filled-out City tax registration form:</b> enclosed you will find the template
    – please send the form back to customer.service@visionapartments.com<br><br>"""
    no_id_p_check = ''
    exit_commands = ('quit', 'close', 'pause', 'off', 'exit',
                     'nothing', 'stop')
    def __init__(self, number=''):
        self.provide_your_name()
        self.mail_count = number
        if self.mail_count == '':
            self.mail_count = int(input('Please provide the number of e-mails to send: ').lower())
            self.check_exit(self.mail_count)
            self.mail_list = [f'{os.getcwd()}\\bookings\{file}' for file in os.listdir('bookings') if file[-4:] == '.msg']

        while self.mail_count > 0:
            self.get_details()
        print("All e-mails sent")
        self.change_file_names()

    def change_file_names(self):
        for name, bnum in self.bookings_names:
            new_name = r'C:\CS_chat_bot\bookings\BOOKING - ' + f'{bnum} Confirmation.msg'
            os.rename(name, new_name)

    def details_forwarded(self, msg):
        c_i = re.findall("Check-in:.*\s*(.+)\n", msg)
        self.c_i = c_i[0].strip()
        c_o = re.findall("Check-Out:.*\s*(.+)\n", msg)
        self.c_o = c_o[0].strip()
        guest_name = re.findall("Guest name:.*\s*(.+)\n", msg)
        self.guest_name = guest_name[0].strip()
        try:
            b_num = re.findall("Reservation no:.*\s*([0-9]+)\s", msg)
            self.b_num = b_num[0].strip()
            email_to = re.findall("Guest email:.*\s*(\S+@\S+)\s*<mailto", msg)
            self.client_address = email_to[0].strip()
        except IndexError:
            print(f"Please check the file ({self.latest_file}), it might be an Airbnb/Amadeus/etc. reservation.")
            self.mail_count -= 1
            return

    def details_no_forwarded(self, msg):
        c_i = re.findall("Check-in:\t(.+)\t", msg)
        self.c_i = c_i[0]
        c_o = re.findall("Check-Out: \t(.+)\t", msg)
        self.c_o = c_o[0]
        guest_name = re.findall("Guest name.+\t(.+)\t", msg)
        self.guest_name = guest_name[0]
        try:
            b_num = re.findall("Reservation no.+\t([0-9]+)", msg)
            self.b_num = b_num[0]
            email_to = re.findall("Guest.+\t(.+@.+) ", msg)
            self.client_address = email_to[0]
        except IndexError:
            print(f"Please check the file ({self.latest_file}), it might be an Airbnb/Amadeus/etc. reservation.")
            self.mail_count -= 1
            raise IndexError

    def details_siteminder(self, msg):
        c_i = re.findall("Check In Date:\s*(.+)\s*", msg)
        self.c_i = c_i[0].strip()
        c_o = re.findall("Check Out Date:\s*(.+)\s*", msg)
        self.c_o = c_o[0].strip()
        guest_name = re.findall("New Reservation\s*(.+)\s*", msg)
        self.guest_name = guest_name[0].strip()
        print(self.c_i, self.c_o, self.guest_name)
        try:
            b_num = re.findall("Booking Confirmation Id:\s*([0-9]+)", msg)
            self.b_num = b_num[0]
            email_to = re.findall("Booker Email:\s*(.+@.+)\s*(?:<mailto)?", msg)
            self.client_address = email_to[0]
            print(self.b_num, self.client_address)
        except IndexError:
            print(f"Please check the file ({self.latest_file}), it might be an Airbnb/Amadeus/etc. reservation.")
            self.mail_count -= 1
            raise IndexError

    def get_details(self):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        file_name = self.mail_list[self.mail_count-1]
        file = outlook.OpenSharedItem(file_name)
        self.latest_file = file_name
        msg = file.Body
        current_count = self.mail_count
        self.check_city(msg)
        print(self.city)
        try:
            if "From: MAGARENTAL" in msg:
                self.details_forwarded(msg)
            elif "support@siteminder.com" in msg:
                self.details_siteminder(msg)
            else:
                self.details_no_forwarded(msg)
        except:
            return

        b_num = input(f'Please provide booking number from VMS for the reservation {self.b_num} (optional)\n ')
        self.check_exit(b_num.lower())
        if b_num == 'skip':
            print(f'Skipping reservation number {self.b_num}')
            self.mail_count -= 1
            return
        if len(b_num) > 6:
            self.b_num = b_num
        self.expedia = self.expedia_check(msg)
        self.id_payment_check(msg)
        self.send_email()
        print(f"Check-in info for booking number {self.b_num} sent")
        self.bookings_names.append((file_name, self.b_num))
        self.mail_count -= 1

    def provide_your_name(self):
        v_name = input('Please, provide your name as at the start of your e-mail (i.e. "Vadym Kulish" -> vkulish): ').lower()
        self.check_exit(v_name)
        df = pd.read_excel('list_emp.xlsx')
        df = df.set_index('init')
        self.name = df.loc[v_name, 'full_name']
        self.title = df.loc[v_name, 'position']
        self.number_local = df.loc[v_name, 'pl_num']
        self.number_foreign = df.loc[v_name, 'sw_num']
        self.my_address = v_name + '@visionapartments.com'
        self.set_signature()

    def check_city(self, message):
        if "VA Frankfurt" in message:
            self.city = "Frankfurt"
            self.p_num = '+49 69 299 170 21'
            self.card = ''

        elif "VA Berlin" in message:
            self.city = "Berlin"
            self.p_num = "+49 30 31 87 67 86"
            self.card = ''

        elif "VA Basel" in message:
            self.city = "Basel"
            self.p_num = "+41 44 248 34 34"
            self.card = ''

        elif "Rue Caroline" in message:
            self.city = "Lausanne Rue Caroline"
            self.p_num = "+41 21 323 96 19"
            self.card = ''

        elif "Chemin des" in message:
            self.city = "Lausanne Chemin des Epinettes"
            self.p_num = "+41 21 323 96 19"
            self.card = ''

        elif "St. Sulpice" in message:
            self.city = "Lausanne Saint-Sulpice"
            self.p_num = "+41 21 323 96 19"
            self.card = ''

        elif "VA Lucerne" in message:
            self.city = "Lucerne"
            self.p_num = "+41 41 508 70 98"
            self.card = self.lucerne_card

        elif "VA Vevey" in message:
            self.city = "Vevey Hotel de Famille"
            self.p_num = "+41 21 510 22 70"
            self.card = self.vevey_card

        elif "VA Vienna" in message:
            self.city = "Vienna"
            self.p_num = "+43 1 229 74 40"
            self.card = ''

        elif "VA Zug" in message:
            self.city = "Zug"
            self.p_num = "+41 41 511 03 56"
            self.card = ''

        elif "VA Zurich" in message:
            self.city = "Zurich"
            self.p_num = "+41 44 248 34 34"
            self.card = ''

        elif "support@siteminder" and 'Vision Nauenstrasse 55' in message:
            self.city = "Basel"
            self.p_num = "+41 44 248 34 34"
            self.card = ''

        else:
            print("Cannot find the city. Try again.")
            raise IndexError

    def expedia_check(self, message):
        if "EXPEDIA" in message:
            return True
        return False

    def id_payment_check(self, message):
        if_id = input(f'Do we have an ID from {self.guest_name}?(yes/no) ').lower()
        self.check_exit(if_id)
        if 'yes' not in if_id:
            self.no_id_p_check = self.requested + self.no_id
        if "received a virtual credit card" not in message:
            if self.expedia:
                reply = input(f'Is there a VCC for this Expedia reservation {self.b_num} ?(yes/no) ').lower()
                self.check_exit(reply)
                if 'yes' in reply:
                    return
                if if_id == 'yes':
                    self.no_id_p_check += self.requested +  self.no_p
                else:
                    self.no_id_p_check += self.no_p
                print(self.b_num, 'No VCC, please send the invoice with a link')
            elif self.city != 'Basel':
                if if_id == 'yes':
                    self.no_id_p_check += self.requested + self.no_p
                else:
                    self.no_id_p_check += self.no_p
                print(self.b_num, 'No VCC, please send the invoice with a link')

    def set_signature(self):
        files = []
        for num in range(6):
            f = open(f'file{num}.txt', 'r')
            text = f.read()
            files.append(text)
            f.close()
        self.signature = '''<p><span style="font-family:Arial,Helvetica,sans-serif">
        <span style="font-size:10pt"><strong><span style="font-size:10.0pt">
        <span style="color:black">'''+ f'{self.name}' '''</span></span></strong></span><br />
        <span style="font-size:10pt"><span style="font-size:10.0pt"><span style="color:black">
        ''' + f'{self.title}' + '''&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        &nbsp;&nbsp;&nbsp; </span></span></span></span></p><p><span style="font-family:Arial,Helvetica,sans-serif">
        <span style="font-size:10pt"><strong><span style="font-size:10.0pt"><span style="color:black">
        Vision Warsaw Sp. z o.o. </span></span></strong></span><br />
        <span style="font-size:10pt"><strong><span style="font-size:10.0pt"><span style="color:black">VISIONAPARTMENTS
        </span></span></strong></span><br /><span style="font-size:10pt"><span style="font-size:10.0pt">
        <span style="color:black">Al. Jerozolimskie 81 I PL-02-001 Warsaw</span></span></span><br />
        <span style="font-size:10pt"><span style="font-size:10.0pt"><span style="color:black">
        T ''' + f'{self.number_local}' + ' I T ' + f'{self.number_foreign}' '''+ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span></span></p><p><span style="font-family:Arial,Helvetica,
        sans-serif"><span style="font-size:10pt"><span style="color:black"><a href="mailto:''' + f'{self.my_address}' + '''"
        style="color:#0563c1; text-decoration:underline"><strong><span style="font-size:10.0pt">
        <span style="color:#4ebcbd">''' + f'{self.my_address}' + '''</span></span></strong></a></span></span><br />
        <span style="font-size:10pt"><span style="color:black"><a href="https://www.visionapartments.com/"
        style="color:#0563c1; text-decoration:underline"><strong><span style="font-size:10.0pt">
        <span style="color:#4ebcbd">visionapartments.com</span></span></strong></a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        &nbsp;&nbsp;&nbsp;</span><span style="font-size:10pt"><span style="font-size:10.0pt"><span style="color:black">
        &nbsp; &nbsp; &nbsp;&nbsp;</span></span></span></span></p><p><span style="font-family:Arial,Helvetica,sans-serif">
        <a href="https://blog.visionapartments.com/"><img src="'''+ files[0] +'''" style="height:25px;
        width:25px" /></a>&nbsp<a href="https://www.facebook.com/FurnishedApartmentsForRent"><img src="'''+ files[1] +'''"
        style="height:25px; width:25px" /></a>&nbsp<a href="https://instagram.com/visionapartments"><img src="'''+ files[2] +'''"
        style="height:25px; width:25px" /></a>&nbsp<a href="https://twitter.com/visionapartment"><img src="'''+ files[3] +'''"
        style="height:25px; width:25px" /></a>&nbsp<a href="https://www.linkedin.com/company/visionapartments"><img src="'''+ files[4] +'''"
        style="height:25px; width:25px" /></a></span></p><table cellspacing="0" class="NormaleTabelle" style="border-collapse:collapse">
        <tbody><tr><td style="vertical-align:top"><p><span style="font-family:Arial,Helvetica,sans-serif"><a href="https://visionapartments.com/">
        <img src="'''+ files[5] +'''" style="height:77px; width:600px" /></a>
        </span></p></td></tr></tbody></table><p><span style="font-family:Arial,Helvetica,sans-serif"><span style=
        "font-size:10pt"><span style="font-size:10.0pt"><span style="color:gray">The controller of your personal data is
        VISIONAPARTMENTS.</span></span> <span style="color:gray"><a href="https://visionapartments.com/en-US/Company/
        Privacy-policy.aspx" style="color:#0563c1; text-decoration:underline"><span style="font-size:10.0pt"><span style=
        "color:#4ebcbd">Read more</span></span></a></span></span></span></p>
'''

    def send_email(self):
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.BCC = self.my_address
        mail.To = self.client_address
        mail.Subject = 'VISIONAPARTMENTS ' + self.city + " / " + self.b_num + " / Check-in information"
        if self.city != 'Zurich':
            info_attach = self.city + "_arrival_tips.pdf"
            info_path = pathlib.Path(info_attach)
            info_abs = str(info_path.absolute())
            mail.Attachments.Add(info_abs)
        if self.city == 'Frankfurt' or self.city == 'Berlin':
            self.no_id_p_check += self.no_f
            form_attach = self.city + "_TOURISMUSBEITRAG_MELDESCHEIN.pdf"
            form_path = pathlib.Path(form_attach)
            form_abs = str(form_path.absolute())
            mail.Attachments.Add(form_abs)
        mail.HTMLBody = '''
            <span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:black'>
            <b>Dear ''' + self.guest_name + '''</b>,<br><br>
            Thank you very much for your reservation.<br><br>
            We are looking forward to welcoming you
            <b>from ''' + self.c_i + ' to ' + self.c_o + '</b> to VISIONAPARTMENTS ' + self.city + '''.<br>
            To grant you access to your apartment, we will provide you with an access code
            on the day of your arrival.<br><br>
            ''' + self.no_id_p_check + '''
            Check in:     possible from 03:00 PM<br>
            Check out:    possible until 10:00 AM<br><br>
            Internet Access:<br><br>
            Network:    VISIONAPARMENTS <br>
            Password:    Wlan4Guest!! <br><br>

            Additionally, we kindly ask you to check the apartment inventory within the
            first three days after your arrival and to inform us about any missing items or
            damages within that time range. On the backside of your entrance door, you will
            find an inventory list which will assist you with this process. Thank you very
            much in advance for your cooperation!<br><br>
            ''' + self.card + '''
            Please do not hesitate to contact our Customer Service team at any time on the
            following number: ''' + self.p_num + '''. Please await the voice message and follow
            the instructions.<br><br>
            Warm regards and a wonderful journey to VISIONAPARTMENTS,<br><br></span>
            ''' + self.signature
        mail.SentOnBehalfOfName = "customer.service@visionapartments.com"
        mail.Send()

    def check_exit(self, reply):
        if reply in self.exit_commands:
            self.change_file_names()
            raise IndexError
