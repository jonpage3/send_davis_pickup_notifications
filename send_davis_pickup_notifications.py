import os, xlrd, codecs, datetime, smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import configparser
config = configparser.ConfigParser()
config.read('send_davis_pickup_notifications.cfg')
workdir = config['GENERAL']['workdir']
importfile = os.path.join(workdir, config['GENERAL']['import_file'])
preview_file = os.path.join(workdir, config['GENERAL']['preview_file'])
debugger_email = config['GENERAL']['debugger_email']
email_template = config['TEMPLATE']['template'].strip()


def ingest_checkouts_export():

    checkouts = {}
    currentrow = 1
    while currentrow <= maxrows:
        this_checkout = {}
        for field in ['barcode','title','volume','due_date','name','email','PID']:
            this_checkout[field] = str(readexcel(field,currentrow))

        pid = this_checkout['PID']
        if pid not in checkouts:
            checkouts[pid] = {'name':this_checkout['name'],
                              'email': this_checkout['email'],
                              'checkout_list' : []}

        #if len(this_checkout['title']) >= 68:
            #this_checkout['title'] = this_checkout['title'][:68] + '...'
        book_title = this_checkout['title'] + " " + this_checkout['volume']
        checkouts[pid]['checkout_list'].append((book_title,this_checkout['due_date']))

        currentrow += 1

    return checkouts


def readexcel(field,rowx):
    if field == 'due_date':
        return f"{(datetime.datetime(*xlrd.xldate_as_tuple(sheet.row_values(rowx)[headings.index(headings_map[field])], bdatemode))):%m-%d-%Y}"
    return sheet.row_values(rowx)[headings.index(headings_map[field])]


def compose_emails(send=False, send_emails_to_patrons=False):
    sent_count = 0
    if send:
        mailserver = smtplib.SMTP('relay.unc.edu', 25)
        mailserver.starttls()
        mailserver.ehlo()
    else:
        try:
            os.remove(preview_file)
        except FileNotFoundError:
            print('Notice: no preview file to delete')

    for patron in patrons:
        itemcount = len(checkouts[patron]['checkout_list'])
        if itemcount == 1:
            itemcount = str(itemcount) + ' item you requested was'
            itemshort = '1 item'
        else:
            if itemcount > 1:
                itemcount = str(itemcount) + ' items you requested are'
                itemshort = str(itemcount).replace(' you requested are', '')
            else:
                print(patron)
                raise RuntimeError('This patron has < 1 requests')

        requestblock = ''
        for checkout in sorted(checkouts[patron]['checkout_list']):
            requestblock += '    ' + checkout[0] + '\n' + '    ' + '\t' + 'Due Date: ' + checkout[1] + '\n'


        requestblock = requestblock.rstrip()

        msg = MIMEMultipart()
        sender = 'Davis Circulation <daviscirc@listserv.unc.edu>'
        msg['From'] = sender

        if not send or send_emails_to_patrons:
            recipient = checkouts[patron]['email']
        else:
            recipient = debugger_email

        msg['To'] = recipient
        msg['Subject'] = 'Library Notice : ' + itemshort + ' Ready for Pick Up'
        message = email_template.replace('{itemcount}', itemcount).replace('{request_block}', requestblock)
        if len(message) < 250:
            print(message)
            print('Length:', len(message))
            raise RuntimeError('This message seems too short to be valid')
        #if len(message.split('\n')) < 18 + len(requests[patron]['req_list']):
            #print(message)
            #print('Lines:', len(message.split('\n')))
            #raise RuntimeError('This message has too few lines to be valid')
        msg.attach(MIMEText(message))

        if send:
            if not previewed:
                print('emails have not been previewed.')
                return False
            mailserver.sendmail(sender, recipient, msg.as_string())
            print('sent email to', recipient)
            sent_count += 1
        else:
            with codecs.open(preview_file, 'a', 'utf-8') as (ofile):
                ofile.write('\nFrom: ' + msg['From'] + '\n')
                ofile.write('To: ' + msg['To'] + '\n')
                ofile.write('Subject: ' + msg['Subject'] + '\n')
                ofile.write(message[:-390] + '...')
            sent_count +=1

    email_list = [checkouts[patron]['email'] for patron in patrons]
    request_list = []
    for patron in checkouts:
        for item in checkouts[patron]['checkout_list']:
            request_list.append(item)

    if sent_count != len(patrons):
        print('generated more emails (', sent_count, ') then patrons (', len(patrons), ')')
        return False
    if maxrows != len(request_list):
        print(maxrows, 'rows on spreadsheet does not match', len(request_list), 'requests')
        return False
    if len(patrons) != len(email_list):
        print('number of patrons (', len(patrons), ') does not match number of email addresses (', len(email_list),')')
        return False
    if len(patrons) + len(email_list) + len(request_list) != len(set(patrons + email_list + request_list)):
        print('there are duplicate patrons or emails or TNs')
        return False
    if send:
        mailserver.quit()
        print('sent', sent_count, 'emails to', len(patrons), 'patrons')
    else:
        print('Found', sent_count, 'emails for', len(patrons), 'patrons for', len(request_list), 'checkouts', '\nExported from Sierra at', datetime.datetime.fromtimestamp(sierra_export_time).strftime('%T'))

        os.startfile(preview_file)
        looksright = input('Does this look correct?(y/n):')
        if looksright == 'y' or looksright == 'yes':
            return True
        return False

def preview_emails():
    if compose_emails(send=False):
        return True

def emails_to_debugger():
    compose_emails(send=True, send_emails_to_patrons=False)

def emails_to_patron():
    compose_emails(send=True, send_emails_to_patrons=True)


headings_map = {'barcode' : 'BARCODE','title':'TITLE','volume':'VOLUME','due_date':'DUE DATE',
                'name':'PATRN NAME','email':'EMAIL ADDR','PID':'P BARCODE'}

if __name__ == '__main__':
    sierra_export_time = os.path.getmtime(importfile)
    if (datetime.datetime.now() - datetime.datetime.fromtimestamp(sierra_export_time)).total_seconds() > 18000:
        raise RuntimeError('ILLiad export is too old')
    print('opening spreadsheet')
    book = xlrd.open_workbook(filename=importfile)
    sheet = book.sheet_by_name('Sheet')
    headings = sheet.row_values(0)
    bdatemode = book.datemode
    print('Ignore a warning above about file size ... not 512 ...\n')
    maxrows = sheet.nrows - 1
    maxcols = sheet.ncols
    checkouts = ingest_checkouts_export()
    #print(checkouts)
    patrons = sorted([patron for patron in checkouts])
    #print(patrons)
    previewed = preview_emails()
    if previewed:
        print('\nTo send emails to patrons, type EMAIL PATRONS')
        emailpatrons = input('type it:')
        if emailpatrons == 'EMAIL PATRONS':
            print('\n')
            emails_to_patron()
            print('Done!')
        else:
            print('okay, should emails be sent to debugger instead?')
            emaildebugger = input('send to debugger (y/n):')
            if emaildebugger == 'y' or emaildebugger == 'yes':
                emails_to_debugger()
                print('emailed ' + debugger_email)
        a = input('\nPress [enter] to exit.')


