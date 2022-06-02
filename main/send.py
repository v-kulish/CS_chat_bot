import win32com.client
import re
import pathlib
import pandas as pd
import os
import openpyxl


class SendMail:
    client_address = ''
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
    no_id = """•    Copy of the ID (*BOTH sides) or Passport:</b> 
    please send it to customer.service@visionapartments.com<br><br>"""
    no_p = "•    <b>Payment:</b> you will receive a separate e-mail with you invoice<br><br>"
    no_f = """•    <b>A filled-out City tax registration form:</b> enclosed you will find the template 
    – please send the form back to customer.service@visionapartments.com<br><br>"""
    no_id_p_check = ''
    def __init__(self, number=''):
        self.provide_your_name()
        self.mail_count = number
        if self.mail_count == '':
            self.mail_count = int(input('Please provide the number of e-mails to send: '))
        while self.mail_count > 0:
            self.get_details()
        print("All e-mails sent")

    def get_details(self):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        old_name = r"C:\CS_chat_bot\bookings\booking " + f'({self.mail_count})' + '.msg'
        file = outlook.OpenSharedItem(old_name)
        msg = file.Body
        self.check_city(msg)
        c_i = re.findall("Check-in:\t(.+)\t", msg)
        self.c_i = c_i[0]
        c_o = re.findall("Check-Out: \t(.+)\t", msg)
        self.c_o = c_o[0]
        guest_name = re.findall("Guest name.+\t(.+)\t", msg)
        self.guest_name = guest_name[0]
        try:
            b_num = re.findall("Reservation no.+\t([0-9]+)", msg)
            self.b_num = b_num[0]
        except IndexError:
            print(f"Please check the file booking ({self.mail_count}), it might be an Airbnb reservation. Closing the program now")
            quit()
        self.b_num = input(f'Please provide booking number from VMS for the reservation {self.b_num} (optional)')
        email_to = re.findall("Guest.+\t(.+.com)", msg)
        self.client_address = email_to[0]
        self.expedia = self.expedia_check(msg)
        self.id_payment_check(msg)
        self.send_email()
        print(f"Check-in info for booking number {self.b_num} sent")
        self.mail_count -= 1

    def provide_your_name(self):
        v_name = input('Please, provide your name as at the start of your e-mail (i.e. "Vadym Kulish" -> vkulish): ').lower()
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

        else:
            print("Cannot find the city. Try again.")
            quit()

    def expedia_check(self, message):
        if "EXPEDIA" in message:
            return True
        return False

    def id_payment_check(self, message):

        if "You have received a virtual credit card" not in message:
            if self.expedia:
                reply = input(f'Is there a VCC for this Expedia reservation{self.b_num} ?(yes/no) ').lower()
                if 'yes' in reply:
                    return
                self.no_id_p_check += self.no_p
                print(self.b_num, 'No VCC, please send the invoice with a link')
            if self.city != 'Basel':
                self.no_id_p_check += self.no_p
                print(self.b_num, 'No VCC, please send the invoice with a link')

    def set_signature(self):
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
        <img src="https://ckeditor.com/apps/ckfinder/userfiles/files/image-20220327012240-1.png" style="height:25px; 
        width:25px" /><img src="https://ckeditor.com/apps/ckfinder/userfiles/files/image-20220327012240-2.png" 
        style="height:25px; width:8px" /><img src="https://ckeditor.com/apps/ckfinder/userfiles/files/image-20220327012240-3.png" 
        style="height:25px; width:25px" /><img src="https://ckeditor.com/apps/ckfinder/userfiles/files/image-20220327012240-4.png" 
        style="height:25px; width:8px" /><img src="https://ckeditor.com/apps/ckfinder/userfiles/files/image-20220327012240-5.png" 
        style="height:25px; width:25px" /><img src="https://ckeditor.com/apps/ckfinder/userfiles/files/image-20220327012240-6.png" 
        style="height:25px; width:8px" /><img src="https://ckeditor.com/apps/ckfinder/userfiles/files/image-20220327012240-7.png" 
        style="height:25px; width:25px" /><img src="https://ckeditor.com/apps/ckfinder/userfiles/files/image-20220327012240-8.png" 
        style="height:25px; width:8px" /><img src="https://ckeditor.com/apps/ckfinder/userfiles/files/image-20220327012240-9.png" 
        style="height:25px; width:25px" /></span></p><table cellspacing="0" class="NormaleTabelle" style="border-collapse:collapse">
        <tbody><tr><td style="vertical-align:top"><p><span style="font-family:Arial,Helvetica,sans-serif"><img src=
        "https://ckeditor.com/apps/ckfinder/userfiles/files/image-20220327012240-10.png" style="height:77px; width:600px" />
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
        mail.To = self.my_address
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
            Thank you very much for your reservation.<br>
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
        stop = input('For checking mistakes')
        mail.Send()