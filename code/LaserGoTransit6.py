import PySimpleGUI as sg
import sys
import mysql.connector
import configparser
import pandas as pd
import json
import copy
import requests
import smtplib
from email.mime.text import MIMEText
import logging
import getpass
import os
import csv
from imapclient import IMAPClient
import webbrowser
import shutil
from bs4 import BeautifulSoup


def selectmaster():

    layout = [[sg.Image(r'Laser2CPLogo.png', background_color='#FFFFFF')],

              [
                  sg.Text('Reference:', text_color='#000000', background_color='#FFFFFF'),
                  sg.Input(key='-MASTER-', size=(11, 1), ),
                  sg.Text('                                  ', text_color='#000000', background_color='#FFFFFF'),
                  sg.Button('OK', button_color=('#000000', '#BABABA'), size=(5, 1), bind_return_key=True, pad=(3, 3)),
                  # sg.Button("TAD Management", button_color=('#000000', '#BABABA'), size=(14, 1), bind_return_key=True, pad=(3, 3)),
                  sg.Button('Close', button_color=('#000000', '#BABABA'), size=(5, 1), pad=(3, 3))],
              ]
    window = sg.Window('LaserGo Transit', layout, background_color='#FFFFFF', icon='customs_icon.ico',
                       no_titlebar=False)
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Close':  # if user closes window or clicks cancel
        window.close()
        sys.exit()
    elif event == "TAD Management":
        window.close()
        requesttad()
        selectmaster()
    elif event == 'OK':
        master = values['-MASTER-']
        window.close()
        return master


def requesttad():

    pendingtaddf = pd.read_csv(pendingtadlocation)
    print(pendingtaddf)

    cprefs = pendingtaddf['CPRef'].tolist()

    tadsready = availabletads()

    layout = [
             [sg.Text('Select CP Reference to request TAD:'
                      , text_color='#000000', background_color='#FFFFFF')],

             [sg.Text('Ref:', text_color='#000000', background_color='#FFFFFF'),
              sg.Combo(cprefs, key='-CPREF-', size=(13, 1)),
              sg.Text('', text_color='#000000', background_color='#FFFFFF'),
              sg.Button('OK', button_color=('#000000', '#BABABA'), size=(5, 1), bind_return_key=True, pad=(3, 3)),
              sg.Button('Close', button_color=('#000000', '#BABABA'), size=(5, 1), pad=(3, 3))],

              [sg.Text("TAD's available for download from Channel Ports", text_color='#000000', background_color='#FFFFFF')],

              [
               sg.Listbox(tadsready, text_color='#000000', background_color='#FFFFFF', size=(37, 10))],

              [sg.Button('Open Customs Pro NCTS Portal', button_color=('#000000', '#BABABA'), size=(34, 1), pad=(3, 3))]

              ]

    window = sg.Window('TAD Management', layout, background_color='#FFFFFF', icon='customs_icon.ico',
                       no_titlebar=False)

    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Close':  # if user closes window or clicks cancel
        window.close()
        return
    if event == 'OK':
        try:
            print(values)
            cpref = values['-CPREF-']
            for ref in pendingtaddf.index:
                if cpref == pendingtaddf['CPRef'][ref]:
                    trackingid = str(pendingtaddf['TrackingID'][ref])
                    vehiclereg = str(pendingtaddf['Vehicle Number'][ref])


            subject = "Re: " + trackingid
            body = "We confirm this vehicle is at our site and we need you to request the NCTS movements are released" \
                    " using the Authorised Consignor approval in place.\n" \
                    "\n" \
                    + "CustomsPro Ref: " + cpref + "\n" \
                    + "Vehicle Reg:  " + vehiclereg + "\n" \
                    + "CustomsPro Tracking Number:  " + trackingid + " -- (with all associated LRN's)" + "\n\n\n" \
                    + "Should you have any problems please contact --- " \
                    + userfullname + " (" + userfunction + ") on " + useremail
        except UnboundLocalError:
            sg.popup('Invalid Channel Ports Reference Entered',
                     background_color='#FFFFFF',
                     text_color='#000000', button_color=('#000000', '#BABABA'), icon='customs_icon.ico', title="Error")
        try:
            sendemail(subject, body, useremail, customsproemail)
        except:
            sg.popup('TAD request email failed to send. Please try again....',
                     background_color='#FFFFFF',
                     text_color='#000000', button_color=('#000000', '#BABABA'), icon='customs_icon.ico', title="Failed")
            window.close()
            return
        pendingtaddf.drop(pendingtaddf[pendingtaddf.CPRef == cpref].index, inplace=True)
        pendingtaddf.to_csv(pendingtadlocation, index=False)
        logentry = 'TAD Request - ' + cpref + ' --- ' + trackingid + ' - ' + currentuser
        logging.info(logentry)
        window.close()
        return

    if event == "Open Customs Pro NCTS Portal":
        webbrowser.open('https://www.customspro.net/ncts-shipments-outuk')
        window.close()
        return


def availabletads():

    server = IMAPClient('outlook.office365.com', use_uid=True)
    server.login('customspro@laserint.co.uk', 'Border2021$')

    select_info = server.select_folder('INBOX')
    print("There are %d TAD's available" % select_info[b'EXISTS'])

    messages = server.search(['FROM', customsproemail])

    tads = set()
    for msgid, data in server.fetch(messages, ['ENVELOPE']).items():
        envelope = data[b'ENVELOPE']
        tads.add("%s" % (envelope.subject.decode())[26:])

    print(tads)

    server.logout()
    return(tads)


def findmasterdetails(masterid):
    try:
        masterdetailsdf = pd.DataFrame()

        masterrow = getmysqldata(fclmasterqry, masterid, fclhost, fcluser, fclpassword, fcldb)
        masterdetailsdf = masterdetailsdf.append(masterrow, ignore_index=True)

        masterdetailsdf.columns = [
            'Agents Reference',
            'Destination Port Date',
            'Destination Code',
            'Origin Port Code',
            'Destination Port Code',
            'Tractor Registration',
            'Trailer Container Number',
            'Consignment Type',
            'On Wheels'
        ]

        masterdetailsdf['Origin Port Code'][0] = masterdetailsdf['Origin Port Code'][0][2:]
        masterdetailsdf['Destination Port Code'][0] = masterdetailsdf['Destination Port Code'][0][2:]

        if masterdetailsdf['Origin Port Code'][0] == 'FOL':
            masterdetailsdf['Origin Port Code'][0] = 'DEU'

    except ValueError:
        return masterdetailsdf

    return masterdetailsdf


def getosaddresses(reference):
    osaddresscodesdf = pd.DataFrame()
    osaddresscoderow = getmysqldata(fclosadresscodes, reference, fclhost, fcluser, fclpassword, fcldb)
    osaddresscodesdf = osaddresscodesdf.append(osaddresscoderow, ignore_index=True)

    osaddresscodesdf.columns = [
        'Job Reference',
        'Service Office of Exit',
        'Importer Code',
        'Agent Code'
    ]

    osimporteraddressdf = pd.DataFrame()
    osagentaddressdf = pd.DataFrame()
    osimportercode = [osaddresscodesdf['Importer Code'][0]]
    print(osaddresscoderow)
    osagentercode = [osaddresscodesdf['Agent Code'][0]]

    osimporteraddressrow = getmysqldata(fclnameaddressqry, osimportercode, fclhost, fcluser, fclpassword, fcldb)
    osagentaddressrow = getmysqldata(fclnameaddressqry, osagentercode, fclhost, fcluser, fclpassword, fcldb)

    osimporteraddresdf = osimporteraddressdf.append(osimporteraddressrow, ignore_index=True)
    osagentaddresdf = osagentaddressdf.append(osagentaddressrow, ignore_index=True)

    print(reference)
    print(osaddresscodesdf)
    print(osimporteraddressdf)

    osimporteraddresdf.columns = [
        'Importer Name',
        'Importer Address 1',
        'Importer Address 2',
        'Importer Address 3',
        'Importer Town',
        'Importer Area Prefix',
        'Importer Area Suffix',
        'Importer Country Code',
        'Importer Country Name'
    ]
    try:
        osagentaddresdf.columns = [
            'Agent Name',
            'Agent Address 1',
            'Agent Address 2',
            'Agent Address 3',
            'Agent Town',
            'Agent Area Prefix',
            'Agent Area Suffix',
            'Agent Country Code',
            'Agent Country Name'
        ]
    except ValueError:
        osagentaddressrow = [('None', 'X', 'X', 'X', 'X', 'X', 'X', 'X', 'X')]
        osagentaddresdf = osagentaddressdf.append(osagentaddressrow, ignore_index=True)
        osagentaddresdf.columns = [
            'Agent Name',
            'Agent Address 1',
            'Agent Address 2',
            'Agent Address 3',
            'Agent Town',
            'Agent Area Prefix',
            'Agent Area Suffix',
            'Agent Country Code',
            'Agent Country Name'
        ]

    print(osaddresscodesdf)
    print(osimporteraddresdf)
    print(osagentaddresdf)

    return osimporteraddresdf, osagentaddresdf


def showmasterdetails(masterdf):
    masterref = [masterdf['Agents Reference'][0]]
    print(masterref)

    if [masterdf['Consignment Type'][0]][0] == '2' \
            or [masterdf['Consignment Type'][0]][0] == '3'\
            or [masterdf['Consignment Type'][0]][0] == '7':
        print('Full Load or Single Job')
        jobrefs = [[masterdf['Agents Reference'][0]]]
        fullsingle = True

    elif [masterdf['Consignment Type'][0]][0] == '1':
        print('Groupage')
        jobrefs = getmysqldata(fcljobsqry, masterref, fclhost, fcluser, fclpassword, fcldb)
        fullsingle = False

    if [masterdf['On Wheels'][0]][0] == 'Y':
        onwheels = 'Yes'
    else:
        onwheels = 'No'
    print(onwheels)

    service = [masterdf['Destination Code'][0]]
    joblist = []
    checked = {}
    for t in jobrefs:
        for x in t:
            joblist.append(x)

    transitoverride = {'-OFFICE-': '', '-OVERRIDE_PARTNER-': False}
    print(transitoverride['-OFFICE-'])
    print(transitoverride['-OVERRIDE_PARTNER-'])

    while True:
        try:
            if transitoverride['-OFFICE-'] == '':
                officeofexit = getmysqldata(fclcustofficeqry, service, fclhost, fcluser, fclpassword, fcldb)
                print(officeofexit)
                officeofexit = officeofexit[0][0]
                print(officeofexit)

                officedetails = checkcustomsoffice(officeofexit)

                # officedetails = getmysqldata(fclofficeofdestqry, officeofexit, fclhost, fcluser, fclpassword, fcldb)
                print(officeofexit)
                print(officedetails)
                officename = ' / ' + officedetails[1]

            else:
                officeofexit = transitoverride['-OFFICE-']
                officedetails = checkcustomsoffice(officeofexit)

                # officename = getmysqldata(fclofficeofdestqry, officeofexit, fclhost, fcluser, fclpassword, fcldb)
                officename = ' / ' + officedetails[1]

        except IndexError:
            sg.popup('No Valid Office of Exit against ' + service[0]
                                                        + ' in FCL. You must update this service to proceed',
                     background_color='#FFFFFF',
                     text_color='#000000', button_color=('#000000', '#BABABA'), icon='customs_icon.ico')
            return

        if transitoverride['-OVERRIDE_PARTNER-'] is False:
            transitpartner = agentaddress['Agent Name']
        else:
            transitpartner = importeraddress['Importer Name']

        layout = [[sg.Image(r'Laser2CPLogo.png', background_color='#FFFFFF')],

                  [sg.Text('Ops Ref: ', size=(9, 1), text_color='#000000', background_color='#FFFFFF'),
                   sg.Text(masterdf['Agents Reference'][0], size=(19, 1), text_color='#000000',
                           background_color='#FFFFFF'),
                   sg.Text('Port Date: ', size=(9, 1), text_color='#000000', background_color='#FFFFFF'),
                   sg.Text(masterdf['Destination Port Date'][0], size=(9, 1), text_color='#000000',
                           background_color='#FFFFFF')],

                  [sg.Text('Origin Port: ', size=(9, 1), text_color='#000000', background_color='#FFFFFF'),
                   sg.Text(masterdf['Origin Port Code'][0], size=(19, 1), text_color='#000000',
                           background_color='#FFFFFF'),
                   sg.Text('Dest. Port: ', size=(9, 1), text_color='#000000', background_color='#FFFFFF'),
                   sg.Text(masterdf['Destination Port Code'][0], size=(9, 1), text_color='#000000',
                           background_color='#FFFFFF')],

                  [sg.Text('Tractor Reg: ', size=(9, 1), text_color='#000000', background_color='#FFFFFF'),
                   sg.Text(masterdf['Tractor Registration'][0], size=(19, 1), text_color='#000000',
                           background_color='#FFFFFF'),
                   sg.Text('Trailer No: ', size=(9, 1), text_color='#000000', background_color='#FFFFFF'),
                   sg.Text(masterdf['Trailer Container Number'][0], size=(9, 1), text_color='#000000',
                           background_color='#FFFFFF')],

                  [sg.Text('Service: ', size=(9, 1), text_color='#000000', background_color='#FFFFFF'),
                   sg.Text(service[0], size=(19, 1), text_color='#000000', background_color='#FFFFFF'),
                   sg.Text('On Wheels: ', size=(9, 1), text_color='#000000', background_color='#FFFFFF'),
                   sg.Text(onwheels, size=(9, 1), text_color='#000000', background_color='#FFFFFF')],

                  #[sg.Text('Delivery is a Full Load or a Single Job.', justification='center', size=(50, 1),
                           #text_color='#FF0000', background_color='#FFFFFF')],
                  #[sg.Text('You may edit the Office of Exit if necessary', justification='center', size=(50, 1),
                           #text_color='#FF0000', background_color='#FFFFFF')],

                  [sg.Text('Office of Exit: ' + officeofexit + officename, justification='center', size=(50, 1),
                           text_color='#000000', background_color='#FFFFFF')],
                  [sg.Text('OS Transit Partner: ' + transitpartner[0], justification='center', size=(50, 1),
                           text_color='#000000', background_color='#FFFFFF')],

                  [sg.Text('', text_color='#000000', background_color='#FFFFFF')],

                  [sg.Button('Review / Send', button_color=('#000000', '#BABABA'), key='_SEND_TO_CP_', pad=(3, 3),
                             size=(12, 1)),
                   sg.Button('Select Jobs', button_color=('#000000', '#BABABA'), key='_JOBS_LIST_', pad=(3, 3),
                             size=(12, 1)),
                   sg.Button('Edit', button_color=('#000000', '#BABABA'), key='_CHANGE_OFFICE_OF_EXIT', pad=(3, 3),
                             size=(12, 1)),
                   sg.Button('Cancel', button_color=('#000000', '#BABABA'), pad=(3, 3), size=(10, 1))]]

        window = sg.Window('Summary - ' + masterdf['Agents Reference'][0], layout, icon='customs_icon.ico',
                           background_color='#FFFFFF')
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel':  # if user closes window or clicks cancel
            window.close()
            return

        elif event == '_CHANGE_OFFICE_OF_EXIT':
            transitoverride = overrideofficeofdest(fullsingle)
            window.close()

        elif event == '_SEND_TO_CP_':
            confirm = sg.popup_yes_no('Confirm', 'Are you sure?',
                                      icon='customs_icon.ico',
                                      background_color='#FFFFFF',
                                      text_color='#000000',
                                      button_color=('#000000', '#BABABA'))

            if confirm == 'No':
                window.close()
                return

            elif confirm == 'Yes':
                print(values)
                preadviceprocessing(masterdf, jobrefs, transitoverride)
                window.close()
                return

        elif event == '_JOBS_LIST_':
            print(jobrefs)
            jobrefs, checked = selectjobs(joblist, checked)
            print(jobrefs)
            window.close()


def selectjobs(joblist, checked):
    print(checked)
    sg.SetOptions(element_padding=(0, 1))

    try:
        lines = [[sg.CB('', key=x, background_color='#FFFFFF', text_color='#000000', default=checked[x]),
                  sg.T('{}'.format(x), background_color='#FFFFFF', text_color='#000000'),
                  sg.T(' ----  ' + (getmysqldata(fcljobnameqry, [x], fclhost, fcluser, fclpassword, fcldb))[0][2]
                       + ' ---- ' + (getmysqldata(fcljobnameqry, [x], fclhost, fcluser, fclpassword, fcldb))[0][0],
                       background_color='#FFFFFF',
                       text_color='#000000')] for x in joblist]
    except KeyError:
        lines = [[sg.CB('', key=x, background_color='#FFFFFF', text_color='#000000', default=True),
                  sg.T('{}'.format(x), background_color='#FFFFFF', text_color='#000000'),
                  sg.T(' ----  ' + (getmysqldata(fcljobnameqry, [x], fclhost, fcluser, fclpassword, fcldb))[0][2]
                       + ' ---- ' + (getmysqldata(fcljobnameqry, [x], fclhost, fcluser, fclpassword, fcldb))[0][0],
                       background_color='#FFFFFF',
                       text_color='#000000')] for x in joblist]

    layout = [[sg.T('Select jobs to include in NCTS submission...',
                    background_color='#FFFFFF',
                    text_color='#000000',
                    font=('Any', 12)
                    )],
              [sg.T('If any jobs are de-selected it is your responsibility to '
                    'ensure transit requirements for all jobs are met.',
                    background_color='#FFFFFF',
                    text_color='#000000',
                    font=('Any', 10)
                    )],
              [sg.T(
                  '',
                  background_color='#FFFFFF',
                  text_color='#000000',
                  font=('Any', 10)
              )],
              *lines,
              [sg.OK(button_color=('#000000', '#BABABA'), size=(5, 1), pad=(3, 3))]
              ]

    form = sg.FlexForm(icon='customs_icon.ico', background_color='#FFFFFF', title='LaserGo Transit - Job Summary')

    button, values = form.Layout(layout).Read()

    selectedjobslist = []
    for job in joblist:
        if values[job] == True:
            z = (job,)
            selectedjobslist.append(z)

    # print(selectedjobslist)

    form.close()

    return selectedjobslist, values


def getmysqldata(query, param, host, user, password, db):
    try:
        cnx = mysql.connector.connect(user=user,
                                      password=password,
                                      host=host,
                                      database=db)
        cursor = cnx.cursor(prepared=True)
        cursor.execute(query, param)
        data = cursor.fetchall()
        cursor.close()
        cnx.close()
    except:
        # print(e)
        logentry = 'FCL dB - Unable to connect to FCL dB' + ' - ' + currentuser
        logging.error(logentry)
        sg.popup('Unable to connect to FCL dB',
                 background_color='#FFFFFF',
                 text_color='#000000', button_color=('#000000', '#BABABA'), icon='customs_icon.ico')
        os.execv(sys.executable, ['python'] + sys.argv)
    return data


def getsadhitemdata(jobreflist):
    sadhitemdf = pd.DataFrame()

    for jobref in jobreflist:
        sadhitemrow = getmysqldata(fclsadhitemqry, jobref, fclhost, fcluser, fclpassword, fcldb)
        sadhitemdf = sadhitemdf.append(sadhitemrow, ignore_index=True)

    try:
        sadhitemdf.columns = ['Job Ref',
                              'Commodity Code',
                              'CPC Number',
                              'Description 1',
                              'Description 2',
                              'Description 3',
                              'Description 4',
                              'Description 5',
                              'Description 6',
                              'Description 7',
                              'Description 8',
                              'Gross Weight',
                              'Net Weight',
                              'Commodity Value',
                              ]
    except ValueError:
        try:
            sg.popup('Job Ref: ' + jobref[0] + ' has no associated customs entry.\r\n'
                                               '\r\n'
                                               '            Transmission of master transit cancelled!\r\n',
                     background_color='#FFFFFF',
                     text_color='#000000',
                     button_color=('#000000', '#BABABA'),
                     icon='customs_icon.ico',
                     title='Transmission Cancelled'
                     )
            emptydf = pd.DataFrame()
            errorsubject = "LaserGo Transit Error "
            errorbody = 'Job Ref: ' + jobref[
                0] + ' has no associated customs entry. Transmission of master transit cancelled!\r\n'
            sendemail(errorsubject, errorbody, lasergoemail, useremail)
            return emptydf
        except (ValueError, UnboundLocalError):
            sg.popup('You must select at least one job for NCTS transmission!',
                     background_color='#FFFFFF',
                     text_color='#000000',
                     button_color=('#000000', '#BABABA'),
                     icon='customs_icon.ico',
                     title='Transmission Cancelled'
                     )
            emptydf = pd.DataFrame()
            return emptydf

    sadhitemdf['Gross Weight'] = sadhitemdf['Gross Weight'].astype(float)
    sadhitemdf['Net Weight'] = sadhitemdf['Net Weight'].astype(float)
    sadhitemdf['Commodity Value'] = sadhitemdf['Commodity Value'].astype(float)

    return sadhitemdf


def getsadhjobdata(jobreflist, itemdf):
    sadhjobdf = pd.DataFrame()

    for jobref in jobreflist:
        totalgrossweight = 0
        sadhrowdata = getmysqldata(fclsadhqry, jobref, fclhost, fcluser, fclpassword, fcldb)
        try:
            sadhrowlist = list(sadhrowdata[0])
        except IndexError:
            sg.popup('Job Ref: ' + jobref[0] + ' has no associated customs entry.\r\n'
                                               '\r\n'
                                               '            Transmission of master transit cancelled!\r\n',
                     background_color='#FFFFFF',
                     text_color='#000000',
                     button_color=('#000000', '#BABABA'),
                     icon='customs_icon.ico',
                     title='Transmission Cancelled'
                     )
            emptydf = pd.DataFrame()
            errorsubject = "LaserGo Transit Error "
            errorbody = 'Job Ref: ' + jobref[0] + ' has no associated customs entry.' \
                                                  'Transmission of master transit cancelled!\r\n'
            sendemail(errorsubject, errorbody, lasergoemail, useremail)
            return emptydf

        for x in range(len(itemdf)):
            if itemdf.loc[x, 'Job Ref'] == jobref[0]:
                grossweight = round((itemdf.loc[x, 'Gross Weight']), 3)
                if grossweight == 0:
                    grossweight = 1
                print(grossweight)
                totalgrossweight = round((totalgrossweight + grossweight), 3)

        print(totalgrossweight)
        sadhrowlist.append(totalgrossweight)

        totalnetweight = 0
        for y in range(len(itemdf)):
            if itemdf.loc[y, 'Job Ref'] == jobref[0]:
                netweight = round((itemdf.loc[y, 'Net Weight']), 3)
                if netweight == 0:
                    netweight = 1
                print(netweight)
                totalnetweight = round((totalnetweight + netweight), 3)

        print(totalnetweight)
        sadhrowlist.append(totalnetweight)

        totalvalue = 0
        for z in range(len(itemdf)):
            if itemdf.loc[z, 'Job Ref'] == jobref[0]:
                value = round((itemdf.loc[z, 'Commodity Value']), 2)
                print(value)
                totalvalue = round((totalvalue + value), 2)

        print(totalvalue)
        sadhrowlist.append(totalvalue)

        # Convert list of tuple into list, append office codes and convert back
        i = 29
        while i < 36:
            custoffice = [sadhrowlist[i]]
            print(custoffice)
            officecodedata = getmysqldata(fclcustofficeqry, custoffice, fclhost, fcluser, fclpassword, fcldb)
            if not officecodedata:
                officecode = ''
            else:
                officecode = officecodedata[0][0]
            print(officecode)
            sadhrowlist.append(officecode)
            i = i + 1

        if sadhrowlist[45] is None:
            partnereori = ''
        else:
            partnereori = sadhrowlist[43] + sadhrowlist[45]
        sadhrowlist.append(partnereori)

        sadhrow = [tuple(sadhrowlist)]
        print(sadhrow)

        sadhjobdf = sadhjobdf.append(sadhrow, ignore_index=True)

    print(sadhjobdf)

    sadhjobdf.columns = ['Job Ref',
                         'User',
                         'User First Name',
                         'User Surname',
                         'Terms',
                         'Entry Number',
                         'Entry Processing Unit',
                         'Declaration Date Time',
                         'Country of Origin',
                         'Country of Destination',
                         'MRN Number',
                         'DUCR',
                         'Exporter Code',
                         'Exporter Name',
                         'Exporter Address',
                         'Exporter Post Code Prefix',
                         'Exporter Post Code Suffix',
                         'Exporter Town County',
                         'Exporter Country',
                         'Exporter EORI',
                         'Consignee Code',
                         'Consignee Name',
                         'Consignee Address',
                         'Consignee Post Code Prefix',
                         'Consignee Post Code Suffix',
                         'Consignee Town County',
                         'Consignee Country',
                         'Total Packages',
                         'Service Office of Origin',
                         'Service Office of Exit',
                         'Service Transit Office 1',
                         'Service Transit Office 2',
                         'Service Transit Office 3',
                         'Service Transit Office 4',
                         'Service Transit Office 5',
                         'Service Transit Office 6',
                         'Partner Office Name',
                         'Partner Address 1',
                         'Partner Address 2',
                         'Partner Address 3',
                         'Partner Office Town City',
                         'Partner Office Area Prefix',
                         'Partner Office Area Suffix',
                         'Partner Office Country',
                         'Partner Office Country Name',
                         'Partner EORI Suffix',
                         'Total Gross Weight',
                         'Total Net Weight',
                         'Total Commodity Value',
                         'Service Office of Exit Code',
                         'Service Transit Office 1 Code',
                         'Service Transit Office 2 Code',
                         'Service Transit Office 3 Code',
                         'Service Transit Office 4 Code',
                         'Service Transit Office 5 Code',
                         'Service Transit Office 6 Code',
                         'Partner EORI'
                         ]
    return sadhjobdf


def checkcustomsoffice(officeid):

    url = "https://www.tariffnumber.com/offices/" + officeid

    resp = requests.get(url)

    print(resp.status_code)

    soup = BeautifulSoup(resp.text, features="lxml")

    try:
        tag = soup.h1
        print("The Customs office code " + officeid + " is for " + tag.text.strip())
        cofficename = tag.text.strip()
        cofficecode = officeid
        if cofficename.startswith('Search for customs offices all over Europe') \
                or cofficename.startswith('Nothing found'):
            raise AttributeError
        else:
            cofficelist = [cofficecode, cofficename]
    except AttributeError:
        print("No Customs office found for code " + officeid)
        cofficelist = []

    return cofficelist


def overrideofficeofdest(fullsingle):

    layout = [[sg.T('Enter alternative Office of Exit code (Blank keeps service default):',
                    background_color='#FFFFFF', text_color='#000000', font=('Any', 10))],

              [sg.T('',
                    background_color='#FFFFFF', text_color='#000000', font=('Any', 5))],

              [sg.T('                 Customs Office Code: ',
                    background_color='#FFFFFF', text_color='#000000', font=('Any', 10)),
               sg.Input('', key='-OFFICE-',
                        background_color='#FFFFFF', text_color='#000000', font=('Any', 10), size=(9, 1))],

              [sg.T('',
                    background_color='#FFFFFF', text_color='#000000', font=('Any', 5))],

              [sg.T('Change Transit Partner to Consignee?',
                    background_color='#FFFFFF', text_color='#000000', font=('Any', 10), visible=fullsingle)],

              [sg.CB('Use "' + importeraddress['Importer Name'][0] + '" as Transit Partner? ',
                     background_color='#FFFFFF', text_color='#000000',
                     font=('Any', 10), key='-OVERRIDE_PARTNER-', visible=fullsingle)],


              [sg.T('WARNING \n '
                    'If you change the customs office on a master that requires \n'
                    'multiple offices of exit (e.g. a trailer that goes from PLO onto AUG) \n '
                    'it will overwrite all of them! If in doubt please ask your line manager',
                    background_color='#FFFFFF', text_color='#FF0000', font=('Any', 10), justification='center')],

              [sg.T('',
                    background_color='#FFFFFF', text_color='#000000', font=('Any', 5))],


              [sg.Button('OK', button_color=('#000000', '#BABABA'),
                         key='_OK_', size=(5, 1), bind_return_key=True, pad=(3, 3)),

               sg.Button('Cancel', button_color=('#000000', '#BABABA'), key='_CANCEL_', size=(5, 1), pad=(3, 3))]]

    window = sg.Window('Edit Transit Details', layout, background_color='#FFFFFF', icon='customs_icon.ico',
                       no_titlebar=False)
    event, values = window.read()

    if event == '_OK_':
        print(event)
        print(values)
        inputoffice = values['-OFFICE-']

        customsofficedetails = checkcustomsoffice(inputoffice)

        # officelist = getmysqldata(fclofficeofdestqry, inputoffice, fclhost, fcluser, fclpassword, fcldb)

        try:

            confirm = sg.popup_yes_no('Office Code: ' + customsofficedetails[0]
                                                      + ' // Location: ' + customsofficedetails[1],
                                      background_color='#FFFFFF',
                                      text_color='#000000', button_color=('#000000', '#BABABA'),
                                      icon='customs_icon.ico',
                                      title='Accept Override Office?')
            if confirm == 'No':
                values = {'-OFFICE-': '', '-OVERRIDE_PARTNER-': False}
                window.close()
                return values
            elif confirm == 'Yes':
                window.close()
                return values
        except IndexError:
            if values['-OFFICE-'] == '':
                window.close()
                return values
            sg.popup('You must enter a valid Office of Exit Code',
                     background_color='#FFFFFF',
                     text_color='#000000', button_color=('#000000', '#BABABA'), icon='customs_icon.ico',
                     title='LaserGo Transit')
            values = {'-OFFICE-': '', '-OVERRIDE_PARTNER-': False}
            window.close()
            return values
    elif event == '_CANCEL_':
        values = {'-OFFICE-': '', '-OVERRIDE_PARTNER-': False}
        window.close()
        return values


def buildcpbulknctsjson(masterdf, jobdf, itemdf, custdest, custtrans, transid):
    data = {}
    data['IsBulk'] = 'true'
    data['CustomerReference'] = transid + masterdf.iloc[0]['Agents Reference'] + currentuser.upper() # Master Ref
    data['VehicleNumber'] = masterdf.iloc[0]['Tractor Registration']  # Job Tractor Reg
    data['TrailerNumber'] = masterdf.iloc[0]['Trailer Container Number']  # Job Trailer Number
    data['ExpectedDate'] = str(masterdf.iloc[0]['Destination Port Date'])  # Date of Movement
    data['RouteUKPortCode'] = masterdf.iloc[0]['Origin Port Code']  # Port of Exit
    data['RouteNonUKPortCode'] = masterdf.iloc[0]['Destination Port Code']  # Arrival Port
    data['LRN'] = ''
    #data['UseAuthorisedLocation'] = 'true'
    #data["AuthorisedLocationCustomsIdentity"] = authorisedlocation

    data['BulkConsignor'] = {
            'Name': 'LASER TRANSPORT INTL LTD.',
            'CountryCode': 'GB',
            'CountryName': 'United Kingdom',
            'Postcode': 'CT21 4LR',
            'AddressLine1': 'LYMPNE DISTRIBUTION PARK',
            'AddressLine2': 'LYMPNE',
            'Town': 'HYTHE',
            'EORI': ltieori,
            'PaymentCode': paymentcode,
            'DefermentNumber': ltideferment,
            'IsDirectRepresentation': 'true'
                            }

    data['BulkConsignee'] = {
        'Name': jobdf.iloc[0]['Partner Office Name'],
        'CountryCode': jobdf.iloc[0]['Partner Office Country'],
        'CountryName': jobdf.iloc[0]['Partner Office Country Name'],
        'Postcode': jobdf.iloc[0]['Partner Office Area Prefix'] + ' ' + jobdf.iloc[0]['Partner Office Area Suffix'],
        'AddressLine1': jobdf.iloc[0]['Partner Address 1'],
        'AddressLine2': jobdf.iloc[0]['Partner Address 2'],
        'Town': jobdf.iloc[0]['Partner Office Town City'],
        'EORI': jobdf.iloc[0]['Partner EORI'],
        'IsDirectRepresentation': 'true'
                    }

    data['BulkGuarantor'] = {
        'Name': 'CHANNELPORTS LTD',
        'CountryCode': 'GB',
        'CountryName': 'United Kingdom',
        'Postcode': 'CT21 4BL',
        'AddressLine1': 'FOLKESTONE SERVICES',
        'AddressLine2': 'JUNCTION 11 M20',
        'Town': 'HYTHE',
        'EORI': 'GB683470514000',
        'IsDirectRepresentation': 'true'
                        }

    data['OfficeOfDestinationNCTSCode'] = custdest  # Office of Destination Code
    data['SecondBorderCrossingNCTSCode'] = custtrans  # Only used if Goods leave EU before final destination
    data['AttachmentContracts'] = []
    data['AttachmentContracts'].append({
            'Name': '',  # Not Required for Laser NCTS
            'Attachment': ''  # Not Required for Laser NCTS
                })

    commodities = []

    for ind in jobdf.index:
        if jobdf['Service Office of Exit Code'][ind] == custdest:
            first = True
            for c in itemdf.index:
                if itemdf['Job Ref'][c] == jobdf['Job Ref'][ind]:
                    if first == True:
                        totalpackages = int(jobdf['Total Packages'][ind])
                        totalgrossweight = jobdf['Total Gross Weight'][ind]
                        totalnetweight = jobdf['Total Net Weight'][ind]
                        totalvalue = jobdf['Total Commodity Value'][ind]

                        commodity = {
                            'LRN': '',
                            'EADMRN': jobdf['MRN Number'][ind],
                            'TotalPackages': int(totalpackages),
                            'TotalGrossWeight': totalgrossweight,
                            'TotalNetWeight': totalnetweight,
                            'InvoiceCurrency': 'GBP',
                            'TotalValue': totalvalue,
                            'BulkCommodityCode': itemdf['Commodity Code'][c],
                            'BulkDescriptionofGoods': itemdf['Description 1'][c] + ' '
                                                      + itemdf['Description 2'][c] + ' ' + itemdf['Description 3'][c],
                                    }
                        first = False
                        commodities.append(commodity)

    data['Consignments'] = commodities

    print(data)
    return data


def buildcpsinglenctsjson(masterdf, jobdf, itemdf, custdest, custtrans, transid):
    data = {}
    data['IsBulk'] = 'false'
    data['CustomerReference'] = transid + masterdf['Agents Reference'][0] + currentuser.upper() # Master Ref
    data['VehicleNumber'] = masterdf['Tractor Registration'][0]  # Job Tractor Reg
    data['TrailerNumber'] = masterdf['Trailer Container Number'][0]  # Job Trailer Number
    data['ExpectedDate'] = str(masterdf['Destination Port Date'][0])  # Date of Movement
    data['RouteUKPortCode'] = masterdf['Origin Port Code'][0]  # Port of Exit
    data['RouteNonUKPortCode'] = masterdf['Destination Port Code'][0]  # Arrival Port
    data['OfficeOfDestinationNCTSCode'] = custdest  # Office of Destination Code
    data['SecondBorderCrossingNCTSCode'] = custtrans  # Only used if Goods leave EU before final destination
    #data['UseAuthorisedLocation'] = 'true'
    #data["AuthorisedLocationCustomsIdentity"] = authorisedlocation

    data['AttachmentContracts'] = []
    data['AttachmentContracts'].append({
        'Name': '',  # Not Required for Laser NCTS
        'Attachment': ''  # Not Required for Laser NCTS
    })

    data['Consignments'] = []
    commodities = []

    for ind in jobdf.index:

        if jobdf['Service Office of Exit Code'][ind] == custdest:

            first = True
            for c in itemdf.index:
                if itemdf['Job Ref'][c] == jobdf['Job Ref'][ind]:

                    commodity = {
                        'CommodityCode': itemdf['Commodity Code'][c],
                        'DescriptionofGoods': itemdf['Description 1'][c] + ' ' + itemdf['Description 2'][c] + ' ' +
                                              itemdf['Description 3'][c],
                        'GrossWeight': itemdf['Gross Weight'][c],
                        'NetWeight': itemdf['Net Weight'][c],
                        'Value': itemdf['Commodity Value'][c],
                    }
                    if first == True:
                        commodity['NumberOfPackages'] = int(jobdf['Total Packages'][ind])
                        first = False
                    else:
                        commodity['NumberOfPackages'] = 0
                    commodities.append(commodity)

            comods = copy.deepcopy(commodities)
            commodity.clear()
            commodities.clear()

            data['Consignments'].append({
                'UKTrader': {
                    'Name': 'LASER TRANSPORT INTL LTD.',
                    'CountryCode': 'GB',
                    'CountryName': 'United Kingdom',
                    'Postcode': 'CT21 4LR',
                    'AddressLine1': 'LYMPNE DISTRIBUTION PARK',
                    'AddressLine2': 'LYMPNE',
                    'Town': 'HYTHE',
                    'EORI': ltieori,
                    'PaymentCode': paymentcode,
                    'DefermentNumber': ltideferment,
                    'IsDirectRepresentation': 'true'
                },
                'Partner': {
                    'Name': jobdf['Partner Office Name'][ind],
                    'CountryCode': jobdf['Partner Office Country'][ind],
                    'CountryName': jobdf['Partner Office Country Name'][ind],
                    'Postcode': jobdf['Partner Office Area Prefix'][ind] + ' ' + jobdf['Partner Office Area Suffix'][
                        ind],
                    'AddressLine1': jobdf['Partner Address 1'][ind],
                    'AddressLine2': jobdf['Partner Address 2'][ind],
                    'Town': jobdf['Partner Office Town City'][ind],
                    'EORI': jobdf['Partner EORI'][ind],
                    'IsDirectRepresentation': 'true'
                },
                'Guarantor': {
                    'Name': 'CHANNELPORTS LTD',
                    'CountryCode': 'GB',
                    'CountryName': 'United Kingdom',
                    'Postcode': 'CT21 4BL',
                    'AddressLine1': 'FOLKESTONE SERVICES',
                    'AddressLine2': 'JUNCTION 11 M20',
                    'Town': 'HYTHE',
                    'EORI': 'GB683470514000',
                    'IsDirectRepresentation': 'true'
                },
                'LRN': jobdf['Job Ref'][ind] + transid,
                'EADMRN': jobdf['MRN Number'][ind],
                "CountryCodeOfDestination": jobdf['Partner Office Country'][ind],
                "CountryNameOfDestination": jobdf['Partner Office Country Name'][ind],
                'TotalPackages': int(jobdf['Total Packages'][ind]),
                'TotalGrossWeight': jobdf['Total Gross Weight'][ind],
                'TotalNetWeight': jobdf['Total Net Weight'][ind],
                'InvoiceCurrency': 'GBP',  # Assumed from SADH
                'TotalValue': jobdf['Total Commodity Value'][ind],
                'Commodities': comods
            })
    print(data)

    return data


def sendcpncts(jsonncts, master, transnumber):
    print(jsonncts)

    jsonauth = {
        "Username": cpapiuser,
        "Password": cpapipwd
    }

    print(jsonauth)

    tokenrequest = requests.post(cpgettokenurl, json=jsonauth)
    tokendict = json.loads(tokenrequest.text)

    apitoken = (tokendict['Token'])
    print(apitoken)

    nctssend = requests.post(cpcreatenctsurl, json=jsonncts, headers={"Authorization": "Bearer " + apitoken})
    print(nctssend.text)

    try:
        response = json.loads(nctssend.text)
        trackingnumber = response['TrackingNumber']
        print(trackingnumber)
        subject = 'Transit Data for ' + master[0] + '-' + transnumber + ' sent'
        body = "The Transit data for " + master[0] + '-' + transnumber + " has been sent to Channel Ports \r\n" \
                                                            "\r\n" \
                                                            "##### Tracking Number: " + trackingnumber + " #####\r\n" \
                                                            "\r\n"
        sg.popup('Success ' + master[0] + "-" + transnumber, 'CP Tracking Number: ' + trackingnumber,
                 background_color='#FFFFFF',
                 text_color='#000000',
                 button_color=('#000000', '#BABABA'), icon='customs_icon.ico'
                 )

        sendemail(subject, body, lasergoemail, useremail)
        logentry = "NCTS Sent - " + master[0] + "-" + transnumber + " --- " + trackingnumber + ' --- ' + currentuser
        logging.info(logentry)
        storetadrequestdata(jsonncts, trackingnumber)

    except (KeyError, ValueError):
        subject = 'Transit Data for ' + master[0] + '-' + transnumber + ' failed'
        body = "The Transit data for " + master[0] + '-' + transnumber + " not sent to Channel Ports with error: \r\n" \
                                                                     "\r\n" \
                                                                     "##### " + nctssend.text + " #####\r\n" \
                                                                                                "\r\n" \
                                                                                                "Please ensure data is re-transmitted"

        sg.popup('The attempted LRN: ' + master[0] + '-' + transnumber + ' transit declaration failed', nctssend.text,
                 background_color='#FFFFFF',
                 text_color='#000000', button_color=('#000000', '#BABABA'), icon='customs_icon.ico', title='GoTransit')

        sendemail(subject, body, lasergoemail, useremail)
        logentry = "NCTS Failed - " + master[0] + "-" + transnumber + " --- " + nctssend.text + ' --- ' + currentuser

        logging.error(logentry)
        # todo this is in the except for test only. For live remove fixed tracking number and move to end of try above
        #trackingnumber = 12345678
        #storetadrequestdata(jsonncts, trackingnumber)
        # todo _______________________________________________________________________________________________________
    return


def storetadrequestdata(cpjson, trackingnum):
    isbulk = cpjson['IsBulk']
    if isbulk == 'true':
        lrn = cpjson['LRN']
    else:
        lrn = (cpjson['Consignments'][0])['LRN']

    cpref = cpjson['CustomerReference']
    vehiclenumber = cpjson['VehicleNumber']
    row = [cpref, vehiclenumber, trackingnum, lrn]
    print(row)

    with open(pendingtadlocation, mode='a', newline='') as file:
        lrnwriter = csv.writer(file, delimiter=',')
        lrnwriter.writerow(row)
    return


def sendemail(subject, body, sender, receiver):

    sg.popup_no_buttons('Emailing details to ' + receiver + '...',
                        icon='customs_icon.ico',
                        non_blocking=True,
                        background_color='#FFFFFF',
                        text_color='#000000',
                        title='GoTransit',
                        auto_close=True
                        )
    msge = body
    msg = MIMEText(msge)
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = receiver
    s = smtplib.SMTP(smtpserver)
    try:
        s.send_message(msg)
    except smtplib.SMTPException as e:
        sg.popup('Failed to send email to recipient ' + sender,
                 background_color='#FFFFFF',
                 text_color='#000000', button_color=('#000000', '#BABABA'), icon='customs_icon.ico', title='GoTransit')
        logentry = 'SMTP Error --- ' + e
        logging.error(logentry)
    s.quit()
    return

def reviewscreen(filename):
    while True:
        layout = [
            # [sg.Image(r'Laser2CPLogo.png', background_color='#FFFFFF')],
            [sg.Text('Please review the transit data that will be sent to CustomsPro and then Send...',
                     text_color='#000000', background_color='#FFFFFF')],
            [sg.Text('', text_color='#000000', background_color='#FFFFFF')],
            [sg.Button('Send', button_color=('#000000', '#BABABA'), size=(5, 1), bind_return_key=True,
                       pad=(3, 3)),
             sg.Button("Review", button_color=('#000000', '#BABABA'), size=(5, 1), bind_return_key=True, pad=(3, 3)),
             sg.Button('Cancel', button_color=('#000000', '#BABABA'), size=(5, 1), pad=(3, 3))],
        ]

        window = sg.Window('LaserGo Transit', layout, background_color='#FFFFFF', icon='customs_icon.ico',
                           no_titlebar=False)
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Cancel':  # if user closes window or clicks cancel
            window.close()
            sendyesyno = 0
            return sendyesyno

        elif event == "Review":
            window.close()
            osCommandString = "notepad.exe " + filename
            os.system(osCommandString)

        elif event == 'Send':
            sendyesno = 1
            window.close()
            return sendyesno


def preadviceprocessing(masterdata, jobrefs, override):
    sg.popup_no_buttons('Calculating the required Transit(s) for ' + masterdata['Agents Reference'][0] + '...',
                        icon='customs_icon.ico',
                        non_blocking=True,
                        background_color='#FFFFFF',
                        text_color='#000000',
                        title='GoTransit',
                        auto_close=True)
    masterref = [masterdata['Agents Reference'][0]]
    print(masterref)


    sadhitemdata = getsadhitemdata(jobrefs)
    if sadhitemdata.empty:
        return

    sadhjobdata = getsadhjobdata(jobrefs, sadhitemdata)
    if sadhjobdata.empty:
        return

    print('***************************************************************************************')
    print(masterdata)
    print(sadhitemdata)
    print('****** JOB REFS ******')
    print(jobrefs)
    print(override)

    if override['-OVERRIDE_PARTNER-'] is True:
        sadhjobdata.iloc[0, sadhjobdata.columns.get_loc('Partner Office Name')] = importeraddress['Importer Name'][
            0]
        sadhjobdata.iloc[0, sadhjobdata.columns.get_loc('Partner Address 1')] = \
            importeraddress['Importer Address 1'][0]
        sadhjobdata.iloc[0, sadhjobdata.columns.get_loc('Partner Address 2')] = \
            importeraddress['Importer Address 2'][0]
        sadhjobdata.iloc[0, sadhjobdata.columns.get_loc('Partner Address 3')] = \
            importeraddress['Importer Address 3'][0]
        sadhjobdata.iloc[0, sadhjobdata.columns.get_loc('Partner Office Town City')] = \
            importeraddress['Importer Town'][0]
        sadhjobdata.iloc[0, sadhjobdata.columns.get_loc('Partner Office Area Prefix')] = \
            importeraddress['Importer Area Prefix'][0]
        sadhjobdata.iloc[0, sadhjobdata.columns.get_loc('Partner Office Area Suffix')] = \
            importeraddress['Importer Area Suffix'][0]
        sadhjobdata.iloc[0, sadhjobdata.columns.get_loc('Partner Office Country')] = \
            importeraddress['Importer Country Code'][0]
        sadhjobdata.iloc[0, sadhjobdata.columns.get_loc('Partner Office Country Name')] = \
            importeraddress['Importer Country Name'][0]
        sadhjobdata.iloc[0, sadhjobdata.columns.get_loc('Partner EORI')] = ''

    if override['-OFFICE-'] != '':
        for job2 in sadhjobdata.index:
            sadhjobdata.iloc[job2, sadhjobdata.columns.get_loc('Service Office of Exit Code')] = override['-OFFICE-']

    destoffices = sadhjobdata['Service Office of Exit Code'].unique()
    transoffice1 = sadhjobdata['Service Transit Office 1 Code'].unique()[0]

    print(jobrefs)
    print(sadhjobdata)
    print(destoffices)
    print(transoffice1)

    # Process and send Transit NCTS /
    transnum = 1
    for destoffice in destoffices:
        sg.popup_no_buttons('Sending transit data to Channel Ports...',
                            icon='customs_icon.ico',
                            non_blocking=True,
                            background_color='#FFFFFF',
                            text_color='#000000',
                            title='GoTransit',
                            auto_close=True
                            )
        print(destoffice)
        print(sadhjobdata)

        if masterdata['Consignment Type'][0] == '1':
            isdestoffice = sadhjobdata['Service Office of Exit Code'] == destoffice
            sadhdestjobdata = sadhjobdata[isdestoffice]
            print('--------------------------------------------------------------------------')
            print(sadhdestjobdata)
            cptransitjson = buildcpbulknctsjson(masterdata, sadhdestjobdata,
                                            sadhitemdata, destoffice, transoffice1, str(transnum))
        else:
            cptransitjson = buildcpsinglenctsjson(masterdata, sadhjobdata,
                                                sadhitemdata, destoffice, transoffice1, str(transnum))

        json_data = json.dumps(cptransitjson, indent=4)
        print(json_data)

        datafilename = str(transnum) + masterref[0] + currentuser.upper() + '.json'

        f = open(datafilename, "w")
        f.write(json_data)
        f.close()

        sendtocp = reviewscreen(datafilename)

        if sendtocp == 1:
            sendcpncts(cptransitjson, masterref, str(transnum))
            transnum = transnum + 1
            shutil.move(datafilename, archivelocation + datafilename)
        else:
            print('Cancelled!')
            os.remove(datafilename)


# ------------ MAIN --------------
config = configparser.ConfigParser()
config.read('LaserGoTransit.ini')
fclhost = config['FCL DB']['host']
fcluser = config['FCL DB']['user']
fclpassword = config['FCL DB']['password']
fcldb = config['FCL DB']['db']
smtpserver = config['EMAIL']['server']
lasergoemail = config['EMAIL']['sender']
smtpdomain = config['EMAIL']['receiver']
customsproemail = config['EMAIL']['customsproemail']
cpgettokenurl = config['API']['cpgettokenurl']
cpcreatenctsurl = config['API']['cpcreatenctsurl']
cpapiuser = config['API']['cpapiuser']
cpapipwd = config['API']['cpapipwd']
ltieori = config['STATIC DATA']['lasereori']
ltideferment = config['STATIC DATA']['ltideferment']
paymentcode = config['STATIC DATA']['paymentcode']
authorisedlocation = config['STATIC DATA']['authorisedlocation']
pendingtadlocation = config['STATIC DATA']['pendingtadlocation']
archivelocation = config['STATIC DATA']['archivelocation']
loglocation = config['LOGGING']['location']
loglevel = config['LOGGING']['level']

fcljobsqry = """    SELECT  OPSREF$$
                    FROM    forwardoffice.CONSIGNMENT_ALL_HEADER
                    WHERE   OPSREF$$_MASTER = ? AND COMPANYEXT = 'LTI'"""

fcljobnameqry = """ SELECT  DOM_NAME,
                            ORIGIN_DEST$$_ORIGIN,
                            ORIGIN_DEST$$_DEST
                    FROM    forwardoffice.CONSIGNMENT_ALL_HEADER
                    WHERE   OPSREF$$ = ? AND COMPANYEXT = 'LTI'"""

fcluserqry = """SELECT  forwardoffice.SUB_SYSTEM_USERS.USER$$ As 'User',
		                forwardoffice.SUB_SYSTEM_USERS.JOB_TITLE As 'Function',
                        forwardoffice.SUB_SYSTEM_USERS.CHRISTIAN_NAME As 'First Name',
                        forwardoffice.SUB_SYSTEM_USERS.SURNAME As 'Surname'
                FROM    forwardoffice.SUB_SYSTEM_USERS

                WHERE   forwardoffice.SUB_SYSTEM_USERS.USER$$ = ? AND COMPANYEXT ='LTI'"""

fclsadhqry = """SELECT	forwardoffice.SADH_ALL_HEADER.REF_OVERLAY AS 'Job Ref',
                        forwardoffice.SADH_ALL_HEADER.USER$$ AS 'User',
                        forwardoffice.SUB_SYSTEM_USERS.CHRISTIAN_NAME AS 'User First Name',
                        forwardoffice.SUB_SYSTEM_USERS.SURNAME AS 'User Surname',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.STERMS$$ AS 'Terms',
                        forwardoffice.SADH_ALL_HEADER.ENTRY_NUMBER AS 'Entry Number',
                        forwardoffice.SADH_ALL_HEADER.EPU$$ AS 'Entry Processing Unit',
                        forwardoffice.SADH_ALL_HEADER.ENTRY$DATETIME AS 'Declaration Date Time',
                        forwardoffice.SADH_ALL_HEADER.COUNTRY$$_ENTRY AS 'Country of Origin',
                        forwardoffice.SADH_ALL_HEADER.COUNTRY$$_DEST AS 'Country of Destination',
                        forwardoffice.SADH_ALL_HEADER.SHMRN AS 'MRN Number',
                        forwardoffice.SADH_ALL_HEADER.UCR$$_DECLN AS 'DUCR',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.CLIENT$$_DOM AS 'Exporter Code',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.DOM_NAME AS 'Shipper Name',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.DOM_ADDR_1 AS 'Shipper Address',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.AREA$$_DOM_PFX AS 'Shipper Post Code Prefix',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.DOM_AREA_CODE_SFX AS 'Shipper Post Code Suffix',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.DOM_TOWN_CITY_CNTRY AS 'Shipper Town County',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.COUNTRY$$_ORIGIN AS 'Shipper Country',
                        forwardoffice.SADH_ALL_HEADER.DECLN_TID AS 'Shipper EORI',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.CLIENT$$_OS AS 'Importer Code',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.OS_NAME AS 'Importer Name',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.OS_ADDR_1 AS 'Importer Address',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.AREA$$_OS AS 'Importer Post Code Prefix',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.OS_AREA_CODE_SFX As 'Importer Post Code Suffix',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.OS_TOWN_CITY_CNTRY AS 'Importer Town County',
                        forwardoffice.CONSIGNMENT_ALL_HEADER.COUNTRY$$_DEST AS 'Importer Country',        
                        forwardoffice.SADH_ALL_HEADER.TOT_PKGS AS 'Total Packages',
                        forwardoffice.FORWARDING_AGENTS_SERVICE.ORIGIN_DEST$$_ORIGIN AS 'Service Office Origin',
                        forwardoffice.FORWARDING_AGENTS_SERVICE.ORIGIN_DEST$$_DEST AS 'Service Office of Exit',
                        forwardoffice.FORWARDING_AGENTS_SERVICE.OFFICE_DEST$$_TRANSIT_1 As 'Service Transit Office 1',
                        forwardoffice.FORWARDING_AGENTS_SERVICE.OFFICE_DEST$$_TRANSIT_2 As 'Service Transit Office 2',
                        forwardoffice.FORWARDING_AGENTS_SERVICE.OFFICE_DEST$$_TRANSIT_3 As 'Service Transit Office 3',
                        forwardoffice.FORWARDING_AGENTS_SERVICE.OFFICE_DEST$$_TRANSIT_4 As 'Service Transit Office 4',
                        forwardoffice.FORWARDING_AGENTS_SERVICE.OFFICE_DEST$$_TRANSIT_5 As 'Service Transit Office 5',
                        forwardoffice.FORWARDING_AGENTS_SERVICE.OFFICE_DEST$$_TRANSIT_6 As 'Service Transit Office 6',
                        forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.NAME_ADDRESS_1 As 'Partner Office Name',
                        forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.NAME_ADDRESS_2 As 'Partner Office Address 1',
                        forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.NAME_ADDRESS_3 As 'Partner Office Address 2',
                        forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.NAME_ADDRESS_4 As 'Partner Office Address 3',
                        forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.TOWN_CITY As 'Partner Office Town City',
                        forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.AREA$$ As 'Partner Office Area Prefix',
                        forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.AREA_CODE_SFX As 'Partner Office Area Suffix',
                        forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.COUNTRY$$ As 'Partner Office Country',
                        forwardoffice.GEN_COUNTRIES.COUNTRY_NAME As 'Partner Office Country Name',
                        forwardoffice.CLIENT_DEF_REGIST_AND_APPROVAL.REFERENCE As 'EORI Suffix'

                FROM 		forwardoffice.CONSIGNMENT_ALL_HEADER
                LEFT JOIN 	forwardoffice.SADH_ALL_HEADER ON forwardoffice.SADH_ALL_HEADER.REF_OVERLAY = forwardoffice.CONSIGNMENT_ALL_HEADER.OPSREF$$
                LEFT JOIN	forwardoffice.SUB_SYSTEM_USERS ON forwardoffice.SADH_ALL_HEADER.USER$$ = forwardoffice.SUB_SYSTEM_USERS.USER$$
                LEFT JOIN	forwardoffice.FORWARDING_AGENTS_SERVICE ON forwardoffice.CONSIGNMENT_ALL_HEADER.ORIGIN_DEST$$_ORIGIN = forwardoffice.FORWARDING_AGENTS_SERVICE.ORIGIN_DEST$$_ORIGIN
                AND			forwardoffice.CONSIGNMENT_ALL_HEADER.ORIGIN_DEST$$_DEST = forwardoffice.FORWARDING_AGENTS_SERVICE.ORIGIN_DEST$$_DEST
                LEFT JOIN	forwardoffice.NAMES_AND_ADDRESSES_CLIENTS ON forwardoffice.FORWARDING_AGENTS_SERVICE.CLIENT$$_AGENT = forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.CLIENT_NUMBER
                LEFT JOIN  forwardoffice.GEN_COUNTRIES ON forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.COUNTRY$$ = forwardoffice.GEN_COUNTRIES.COUNTRY$$
                AND         forwardoffice.GEN_COUNTRIES.COMPANYEXT = 'LTI'

                LEFT JOIN	forwardoffice.CLIENT_DEF_REGIST_AND_APPROVAL ON forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.CLIENT_NUMBER = forwardoffice.CLIENT_DEF_REGIST_AND_APPROVAL.CLIENT$$
                AND         forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.COUNTRY$$ = forwardoffice.CLIENT_DEF_REGIST_AND_APPROVAL.COUNTRY$$
                AND			forwardoffice.CLIENT_DEF_REGIST_AND_APPROVAL.REGAPP$$ = 'EORI'

                where forwardoffice.SADH_ALL_HEADER.COMPANYEXT='LTI' 
                and   forwardoffice.SADH_ALL_HEADER.RECORD_TYPE='01' 
                and   forwardoffice.SADH_ALL_HEADER.CUSENTCAT$$='E'
                and   forwardoffice.SADH_ALL_HEADER.COUNTRY$$_ENTRY='GB' 
                and   forwardoffice.SADH_ALL_HEADER.REFERENCE_TYPE=1 
                and   forwardoffice.SADH_ALL_HEADER.REF_OVERLAY= ?
                and   forwardoffice.SADH_ALL_HEADER.SHSTATUS_CODE$$_PRG in (9, 10) 
                and   forwardoffice.SADH_ALL_HEADER.TRAINING_ENTRY !=1"""

fclsadhitemqry = """SELECT 
                        forwardoffice.SADH_ALL_ITEMS.REF_OVERLAY AS 'Job Ref',
                        forwardoffice.SADH_ALL_ITEMS.TARIC_CMDTY_CODE AS 'Commodity Code',
                        forwardoffice.SADH_ALL_ITEMS.CPC AS 'CPC Number',
                        forwardoffice.SADH_ALL_ITEMS.GDS_DESC_ARRAY_1 AS 'Description 1',
                        forwardoffice.SADH_ALL_ITEMS.GDS_DESC_ARRAY_2 AS 'Description 2',
                        forwardoffice.SADH_ALL_ITEMS.GDS_DESC_ARRAY_3 AS 'Description 3',
                        forwardoffice.SADH_ALL_ITEMS.GDS_DESC_ARRAY_4 AS 'Description 4',
                        forwardoffice.SADH_ALL_ITEMS.GDS_DESC_ARRAY_5 AS 'Description 5',
                        forwardoffice.SADH_ALL_ITEMS.GDS_DESC_ARRAY_6 AS 'Description 6',
                        forwardoffice.SADH_ALL_ITEMS.GDS_DESC_ARRAY_7 AS 'Description 7',
                        forwardoffice.SADH_ALL_ITEMS.GDS_DESC_ARRAY_8 AS 'Description 8',
                        forwardoffice.SADH_ALL_ITEMS.ITEM_GROSS_MASS AS 'Gross Weight',
                        forwardoffice.SADH_ALL_ITEMS.ITEM_NETT_MASS AS 'Net Weight',
                        forwardoffice.SADH_ALL_ITEMS.ITEM_STAT_VAL_DC AS 'Commodity Value'
                    FROM
                        forwardoffice.SADH_ALL_ITEMS
                    JOIN
                        forwardoffice.SADH_ALL_HEADER 
                        ON forwardoffice.SADH_ALL_HEADER.COMPANYEXT = forwardoffice.SADH_ALL_ITEMS.COMPANYEXT
                    AND forwardoffice.SADH_ALL_HEADER.RECORD_TYPE = '01'
                    AND forwardoffice.SADH_ALL_HEADER.CUSENTCAT$$ = forwardoffice.SADH_ALL_ITEMS.CUSENTCAT$$
                    AND forwardoffice.SADH_ALL_HEADER.COUNTRY$$_ENTRY = forwardoffice.SADH_ALL_ITEMS.COUNTRY$$_ENTRY
                    AND forwardoffice.SADH_ALL_HEADER.REFERENCE_TYPE = forwardoffice.SADH_ALL_ITEMS.REFERENCE_TYPE
                    AND forwardoffice.SADH_ALL_HEADER.REF_OVERLAY = forwardoffice.SADH_ALL_ITEMS.REF_OVERLAY
                    AND forwardoffice.SADH_ALL_HEADER.REFERENCE_CATEGORY = forwardoffice.SADH_ALL_ITEMS.REFERENCE_CATEGORY
                    AND forwardoffice.SADH_ALL_HEADER.REFERENCE_PART = forwardoffice.SADH_ALL_ITEMS.REFERENCE_PART
                    WHERE
                        forwardoffice.SADH_ALL_ITEMS.COMPANYEXT = 'LTI'
                    AND forwardoffice.SADH_ALL_ITEMS.RECORD_TYPE = '06'
                    AND forwardoffice.SADH_ALL_ITEMS.CUSENTCAT$$ = 'E'
                    AND forwardoffice.SADH_ALL_ITEMS.COUNTRY$$_ENTRY = 'GB'
                    AND forwardoffice.SADH_ALL_ITEMS.REFERENCE_TYPE = 1
                    AND forwardoffice.SADH_ALL_ITEMS.REF_OVERLAY = ?
                    AND forwardoffice.SADH_ALL_HEADER.SHSTATUS_CODE$$_PRG IN (9 , 10, 11)"""

fclcustofficeqry = """  SELECT  SHCUSOFF$$ As 'Customs Office Code'
                        FROM    forwardoffice.GEN_CUSTOMS_OFFICE_CUSDAT
                        WHERE   forwardoffice.GEN_CUSTOMS_OFFICE_CUSDAT.OFFICE_DEST$$ = ?"""

fclservicedestqry = """ SELECT  OFFICE_DEST$$ As 'Service Office Destination'
                        FROM    forwardoffice.GEN_CUSTOMS_OFFICE_CUSDAT
                        WHERE   COMPANYEXT = 'GLB' AND CUSENTCAT$$ = 'E' AND
                                SHCUSOFF$$ = ?"""

fclmasterqry = """SELECT    forwardoffice.CONSIGNMENT_ALL_SHIP_DETAILS.OPSREF$$ As 'Agents Reference',
                            forwardoffice.CONSIGNMENT_ALL_SHIP_DETAILS.DOM_PORT_DATE As 'Activity Date',
                            forwardoffice.CONSIGNMENT_ALL_HEADER.ORIGIN_DEST$$_OVERSEAS As 'Destination Code',
                            forwardoffice.CONSIGNMENT_ALL_SHIP_DETAILS.HOMEPORT$$ As 'Origin Port Code',
                            forwardoffice.CONSIGNMENT_ALL_SHIP_DETAILS.OSEAPORT$$ As 'Destination Port Code',
                            forwardoffice.CONSIGNMENT_ALL_SHIP_DETAILS.TRACTOR_REG As 'Tractor Registration',
                            forwardoffice.CONSIGNMENT_ALL_SHIP_DETAILS.EQUIP_CODE$$_TIR As 'Trailer Contr Nationality',
                            forwardoffice.CONSIGNMENT_ALL_HEADER.CONS_TYPE$$ As 'Consignment Type',
                            forwardoffice.CONSIGNMENT_ALL_SHIP_DETAILS.OS_CARTAGE_ON_WHEELS As 'On Wheels'

                  FROM      forwardoffice.CONSIGNMENT_ALL_SHIP_DETAILS
                  LEFT JOIN forwardoffice.CONSIGNMENT_ALL_HEADER 
                  ON        forwardoffice.CONSIGNMENT_ALL_SHIP_DETAILS.OPSREF$$ = forwardoffice.CONSIGNMENT_ALL_HEADER.OPSREF$$

                  WHERE     forwardoffice.CONSIGNMENT_ALL_SHIP_DETAILS.OPSREF$$ = ?
                  AND       forwardoffice.CONSIGNMENT_ALL_SHIP_DETAILS.COMPANYEXT = 'LTI'"""

fclofficeofdestqry = """SELECT forwardoffice.CUS_OFFICE_OF_EXIT.SHCUSOFF$$ As 'Office of Exit Code',
                               forwardoffice.CUS_OFFICE_OF_EXIT.OFFICE_NAME As 'Office of Exit Location'

                        FROM   forwardoffice.CUS_OFFICE_OF_EXIT

                        WHERE   COMPANYEXT = 'GLB' AND CUSENTCAT$$ = 'E' AND
                                SHCUSOFF$$ = ?"""

fclnameaddressqry = """SELECT 	forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.NAME_ADDRESS_1 As 'Name',
                                forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.NAME_ADDRESS_2 As 'Address 1',
                                forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.NAME_ADDRESS_3 As 'Address 3',
                                forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.NAME_ADDRESS_4 As 'Address 4',
                                forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.TOWN_CITY As 'Town',
                                forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.AREA$$ As 'Area Prefix',
                                forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.AREA_CODE_SFX As 'Area Suffix',
                                forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.COUNTRY$$ As 'Country Code',
                                forwardoffice.GEN_COUNTRIES.COUNTRY_NAME As 'Country Name'      

                        FROM forwardoffice.NAMES_AND_ADDRESSES_CLIENTS
                        LEFT JOIN forwardoffice.GEN_COUNTRIES
                        ON forwardoffice.NAMES_AND_ADDRESSES_CLIENTS.COUNTRY$$ = forwardoffice.GEN_COUNTRIES.COUNTRY$$
                        WHERE CLIENT_NUMBER = ? AND forwardoffice.GEN_COUNTRIES.COMPANYEXT = 'LTI'"""

fclosadresscodes = """  SELECT    forwardoffice.CONSIGNMENT_ALL_HEADER.OPSREF$$ AS 'Job Ref',
                                forwardoffice.FORWARDING_AGENTS_SERVICE.ORIGIN_DEST$$_DEST AS 'Service Office of Exit',
                                forwardoffice.CONSIGNMENT_ALL_HEADER.CLIENT$$_OS As 'Importer Code',
                                forwardoffice.FORWARDING_AGENTS_SERVICE.CLIENT$$_AGENT As 'Agent Code'                   
                        FROM 		forwardoffice.CONSIGNMENT_ALL_HEADER

                        LEFT JOIN	forwardoffice.FORWARDING_AGENTS_SERVICE ON forwardoffice.CONSIGNMENT_ALL_HEADER.ORIGIN_DEST$$_ORIGIN = forwardoffice.FORWARDING_AGENTS_SERVICE.ORIGIN_DEST$$_ORIGIN
                        AND			forwardoffice.CONSIGNMENT_ALL_HEADER.ORIGIN_DEST$$_DEST = forwardoffice.FORWARDING_AGENTS_SERVICE.ORIGIN_DEST$$_DEST                

                        WHERE	forwardoffice.CONSIGNMENT_ALL_HEADER.OPSREF$$ = ?"""

pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

logging.basicConfig(
    filename=loglocation + 'LaserGoTransit.log',
    format='%(asctime)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S',
    level=logging.getLevelName(loglevel)
                    )

# todo -------------- Remove fixed user -------------

currentuser = 'jpo'
# currentuser = getpass.getuser()
userdetails = getmysqldata(fcluserqry, [currentuser], fclhost, fcluser, fclpassword, fcldb)

try:
    userfullname = userdetails[0][2] + ' ' + userdetails[0][3]
    userfunction = userdetails[0][1]
    useremail = currentuser + "@" + smtpdomain
except IndexError:
    sg.popup('You must have a valid FCL account to use this application',
             background_color='#FFFFFF',
             text_color='#000000', button_color=('#000000', '#BABABA'), icon='customs_icon.ico', title="Error")
    sys.exit()

print(currentuser)
print(userfullname)
print(userfunction)
print(useremail)

while True:
    selectedmaster = [selectmaster()]
    print(selectedmaster)
    masterdetails = findmasterdetails(selectedmaster)
    try:
        importeraddress, agentaddress = getosaddresses(selectedmaster)
        print(masterdetails)
        showmasterdetails(masterdetails)
    except ValueError:
        sg.popup('The reference ' + str(selectedmaster[0]) + ' was not found in FCL',
                 background_color='#FFFFFF',
                 text_color='#000000',
                 button_color=('#000000', '#BABABA'),
                 title='GoTransit',
                 icon='customs_icon.ico')
