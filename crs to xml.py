import datetime
import time
import sys
import shutil
import os
import re
from xml.etree import ElementTree
from xml.etree.ElementTree import Element, SubElement
from xml.dom import minidom
import pandas
import pycountry

__author__ = "Timothy Cameron"
__email__ = "tcameron@devtechsys.com"
__date__ = "09-06-2018"
__version__ = "0.1"
date = datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3]+'Z'


def prettify(elem):
    """
    Return a pretty-printed string for the xml elements.
    :param elem: The current XML tree.
    :return reparsed: The pretty-printed string from the XML tree.
    """
    rough_string = ElementTree.tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")


def group_split(crsgroupfile):
    """
    Return lists of separate recipients and the ids for each row.
    :param crsgroupfile: The source file to grab the data from
    :return ids: The main activity identifiers
    """

    actives = list()
    ids = list()
    for i in range(0, len(crsgroupfile.index)):
        try:
            crs_id = str(crsgroupfile["crs_id_number"][i])
        except ValueError:
            crs_id = '000'
        print(crs_id)
        trans_crs_id = re.sub("\D", "", crs_id)
        print(trans_crs_id)
        if trans_crs_id not in actives:
            actives.append(trans_crs_id)
            ids.append([trans_crs_id, i])
        else:
            for each in ids:
                if trans_crs_id == each[0]:
                    each.append(i)
    return ids


def country_convert(input_country):
    # input_countries = [input_country]

    # countries = {}
    # for count in pycountry.countries:
    #     countries[count.name] = count.alpha_2
    # code = [countries.get(country, 'Unknown code') for country in input_countries]
    try:
        code = pycountry.countries.get(name=input_country).alpha_2
    except KeyError:
        try:
            code = pycountry.countries.get(common_name=input_country).alpha_2
        except KeyError:
            try:
                code = pycountry.countries.get(official_name=input_country).alpha_2
            except KeyError:
                code = ''
    return code


def open_files():
    """
    :return crs: The crs file with the main data
    """
    # Prompt user for filename
    # filetoopen = input("What is the name of the crs source file? ")
    filetoopen = 'crs/new_crs1.xlsx'
    # filetoopen = sys.argv[1]
    print('Opening CRS file...')
    # Read the file
    try:
        crs = pandas.read_excel(filetoopen, encoding='utf-8')
    except FileNotFoundError:
        sys.exit("CRS file does not exist.")
    # Output the number of rows
    print('Total rows: {0}'.format(len(crs)))
    # See which headers are available
    print(list(crs))

    return crs


crsfile = open_files()
print("Converting format...")

crsgrouping = group_split(crsfile)
print(crsgrouping)

ver = '2.03'
fasite = 'https://explorer.usaid.gov/'

activities = Element('iati-activities', version=ver,
                     generated_h_datetime=date, xmlns__usg=fasite)

for crsAct in crsgrouping:

    baseLine = crsAct[1]

    activityid = crsAct[0]
    activity = SubElement(activities, 'iati-activity', last_h_updated_h_datetime=date, xml__lang='en', hierarchy='1')

    # A2 is the CRS type
    otheridentifier = SubElement(activity, 'other-identifier', ref=activityid, type='A2')

    title = str(crsfile["project_title"][baseLine])
    description = str(crsfile["description"][baseLine])

    if title != "" and title != 'nan':
        awardtitle = SubElement(activity, 'title')
        narrative = SubElement(awardtitle, 'narrative')
        narrative.text = title
    if description != "" and description != 'nan':
        awarddescription = SubElement(activity, 'description')
        narrative = SubElement(awarddescription, 'narrative')
        narrative.text = description

    startDate = str(crsfile["start_date"][baseLine])
    endDate = str(crsfile["completion_date"][baseLine])

    if startDate != "" and startDate != 'nan':
        activityDate = SubElement(activity, 'activity-date', type='1', iso_h_date=startDate[:10])
    if endDate != "" and endDate != 'nan':
        activityDate = SubElement(activity, 'activity-date', type='3', iso_h_date=endDate[:10])

    # location
    loc = str(crsfile["geographical_target"][baseLine])
    if loc != 'nan':
        location = SubElement(activity, 'location')
        locationname = SubElement(location, 'name')
        narrative = SubElement(locationname, 'narrative')
        narrative.text = loc

    # policy-marker code=1-9

    # 1 gender equality ["gender_equity"]
    try:
        gender = str(int(crsfile["gender_equity"][baseLine]))
    except ValueError:
        gender = ''
    if gender != '' and gender != 'nan':
        genderequality = SubElement(activity, 'policy-marker', code='1', vocabulary='1', significance=gender)
    # 2 aid to environment ["aid_to_environment"]
    try:
        environment = str(int(crsfile["aid_to_environment"][baseLine]))
    except ValueError:
        environment = ''
    if environment != '' and environment != 'nan':
        aidtoenvironment = SubElement(activity, 'policy-marker', code='2', vocabulary='1', significance=environment)
    # 3 pd/gg ["pd_GG"]
    try:
        pdgg = str(int(crsfile["pd_GG"][baseLine]))
    except ValueError:
        pdgg = ''
    if pdgg != '' and pdgg != 'nan':
        pdGG = SubElement(activity, 'policy-marker', code='3', vocabulary='1', significance=pdgg)
    # 4 trade development ["Trade_Development"]
    try:
        trade = str(int(crsfile["Trade_Development"][baseLine]))
    except ValueError:
        trade = ''
    if trade != '' and trade != 'nan':
        tradeDevelopment = SubElement(activity, 'policy-marker', code='4', vocabulary='1', significance=trade)
    # 5 biodiversity ["biodiversity"]
    try:
        bio = str(int(crsfile["biodiversity"][baseLine]))
    except ValueError:
        bio = ''
    if bio != '' and bio != 'nan':
        biodiversity = SubElement(activity, 'policy-marker', code='5', vocabulary='1', significance=bio)
    # 6 climate change mitigation ["mitigation"]
    try:
        mitigation = str(int(crsfile["mitigation"][baseLine]))
    except ValueError:
        mitigation = ''
    if mitigation != '' and mitigation != 'nan':
        climatemitigation = SubElement(activity, 'policy-marker', code='6', vocabulary='1', significance=mitigation)
    # 7 climate change adaptation ["adaptation"]
    try:
        adaptation = str(int(crsfile["adaptation"][baseLine]))
    except ValueError:
        adaptation = ''
    if adaptation != '' and adaptation != 'nan':
        climateadaptation = SubElement(activity, 'policy-marker', code='7', vocabulary='1', significance=adaptation)
    # 8 desertification ["desertification"]
    try:
        desert = str(int(crsfile["desertification"][baseLine]))
    except ValueError:
        desert = ''
    if desert != '' and desert != 'nan':
        desertification = SubElement(activity, 'policy-marker', code='8', vocabulary='1', significance=desert)
    # 9 rmnch ["RMNCH"]
    try:
        rmnch = str(int(crsfile["RMNCH"][baseLine]))
    except ValueError:
        rmnch = ''
    if rmnch != '' and rmnch != 'nan':
        ranch = SubElement(activity, 'policy-marker', code='9', vocabulary='1', significance=rmnch)

    try:
        collab = str(int(crsfile["bi_multi"][baseLine]))
    except ValueError:
        collab = 0
    try:
        flow = str(int(crsfile["flow_type"][baseLine]))
    except ValueError:
        flow = 0
    try:
        finance = str(int(crsfile["finance_type"][baseLine]))
    except ValueError:
        finance = 0
    try:
        aid = str(crsfile["dac_typology"][baseLine])
    except ValueError:
        aid = 0

    if collab != 0:
        collabtype = SubElement(activity, 'collaboration-type', code=collab)
    if flow != 0:
        flowtype = SubElement(activity, 'default-flow-type', code=flow)
    if finance != 0:
        financetype = SubElement(activity, 'default-finance-type', code=finance)
    if aid != 'nan' and aid != '':
        aidtype = SubElement(activity, 'default-aid-type', code=aid)

    # default-aid-type (?)
    # capital-spend (not applicable)

    for trans in crsAct[1:]:

        # try:
        #    usaidaward = str(crsfile["usaid_award_number"][trans])
        #    if usaidaward != "" and usaidaward != 'nan' and "Administrative Costs" not in usaidaward:
        #        identifier = SubElement(activity, 'iati-identifier', ref=usaidaward)
        # except ValueError:
        #    usaidaward = ''

        # Variables that depend on entries
        # If the disbursement has a value, set value to disbursement.
        transaction_code = ''
        try:
            transAmount = float(crsfile["amt_extended"][trans])
            transaction_code = '3'
        except ValueError:
            try:
                transAmount = float(crsfile["commitments"][trans])
                transaction_code = '2'
            except ValueError:
                try:
                    transAmount = float(crsfile["amt_received"][trans])
                    transaction_code = '6'
                except ValueError:
                    try:
                        transAmount = float(crsfile["interest_received"][trans])
                        transaction_code = '6'
                    except ValueError:
                        transAmount = '0'
        valueAmount = str('{0:.2f}'.format(transAmount))

        if valueAmount != 0 and valueAmount != 'nan':
            value_datetime = '2017-01-01'

            transaction = SubElement(activity, 'transaction')
            transaction_type = SubElement(transaction, 'transaction-type',
                                          code=transaction_code)
            transaction_date = SubElement(transaction, 'transaction-date',
                                          iso_h_date=value_datetime)
            value = SubElement(transaction, 'value',
                               value_h_date=value_datetime)
            value.text = valueAmount

            country = country_convert(str(crsfile["recipient_country"][trans]))
            # country = str(crsfile["recipient_country"][baseLine])
            if country != '':
                recipient = SubElement(transaction, 'recipient-country', code=country)

            # sector
            try:
                sector = str(int(crsfile["purpose_code"][trans]))
            except ValueError:
                sector = 0
            if sector != 0:
                sectorcode = SubElement(transaction, 'sector', code=str(sector), percentage='100')

            tiedstatus = ''
            try:
                tiedAmount = float(crsfile["amt_untied"][trans])
                tiedstatus = '5'
            except ValueError:
                try:
                    tiedAmount = float(crsfile["amt_partial"][trans])
                    tiedstatus = '3'
                except ValueError:
                    try:
                        tiedAmount = float(crsfile["amt_tied"][trans])
                        tiedstatus = '4'
                    except ValueError:
                        tiedAmount = 0
            valueAmount = str('{0:.2f}'.format(tiedAmount))

            if valueAmount != 0 and valueAmount != 'nan':
                value_datetime = '2017-01-01'

                tied = SubElement(transaction, 'dac__tied-status', code=tiedstatus)

                tied_date = SubElement(tied, 'dac__value', value_h_date=value_datetime)
                tied_date.text = valueAmount

    # CRS-ADD fields
    crsAdd = SubElement(activity, "crs-add")
    ftcIt = False
    pbaIt = False
    ipIt = False
    afIt = False

    for line in crsAct[1:]:

        try:
            reportyear = int(crsfile["reporting_year"][line])
        except ValueError:
            reportyear = 0
        try:
            comdate = str(crsfile["commitment_date"][line])[0:10]
        except ValueError:
            comdate = 0

        # other-flags code='1' "FTC"
        try:
            ftc = str(int(crsfile["FTC"][line]))
        except ValueError:
            ftc = ''
        # other-flags code='2' "Programme_based_approach"
        try:
            pba = str(int(crsfile["Programme_based_approach"][line]))
        except ValueError:
            pba = ''
        # other-flags code='3' "investment_project"
        try:
            ip = str(int(crsfile["investment_project"][line]))
        except ValueError:
            ip = ''
        # other-flags code='4' "AF"
        try:
            af = str(int(crsfile["AF"][line]))
        except ValueError:
            af = ''

        if ftc != '' and ftc != 'nan' and ftcIt is False:
            ftcAdd = SubElement(crsAdd, 'other-flags', code='1', significance=ftc)
            ftcIt = True
        if pba != '' and pba != 'nan' and pbaIt is False:
            pbaAdd = SubElement(crsAdd, 'other-flags', code='2', significance=pba)
            pbaIt = True
        if ip != '' and ip != 'nan' and ipIt is False:
            ipAdd = SubElement(crsAdd, 'other-flags', code='3', significance=ip)
            ipIt = True
        if af != '' and af != 'nan' and afIt is False:
            afAdd = SubElement(crsAdd, 'other-flags', code='4', significance=af)
            afIt = True

        # DAC Loan Terms block
        # dac:loan-terms rate-1=["Interest_rate"] rate-2=["Second_interest_rate"]
        #   dac:repayment-type code=["Type"]
        #   dac:repayment-plan code=["No_repayments"]
        #   dac:repayment-first-date iso-date=["First_repay_date"]
        #   dac:repayment-final-date iso-date=["Final_repay_date"]
        try:
            rate1 = str(float(crsfile["Interest_rate"][line]))
        except ValueError:
            rate1 = ''
        try:
            rate2 = str(float(crsfile["Second_interest_rate"][line]))
        except ValueError:
            rate2 = ''

        try:
            repaymentType = str(int(crsfile["Type"][line]))
        except ValueError:
            repaymentType = ''
        try:
            repaymentPlan = str(int(crsfile["No_repayments"][line]))
        except ValueError:
            repaymentPlan = ''
        repaymentFirst = str(crsfile["First_repay_date"][line])
        repaymentFinal = str(crsfile["Final_repay_date"][line])

        loanterms = ''
        if (rate1 != 'nan') or (rate2 != 'nan'):
            if rate1 != 'nan' and rate2 != 'nan':
                loanterms = SubElement(crsAdd, 'dac__loan-terms', rate_h_1=rate1, rate_h_2=rate2)
            elif rate1 != 'nan':
                loanterms = SubElement(crsAdd, 'dac__loan-terms', rate_h_1=rate1)
            elif rate2 != 'nan':
                loanterms = SubElement(crsAdd, 'dac__loan-terms', rate_h_2=rate2)
            if repaymentType != 'nan':
                repayType = SubElement(loanterms, 'dac__repayment-type', code=repaymentType)
            if repaymentPlan != 'nan':
                repayPlan = SubElement(loanterms, 'dac__repayment-plan', code=repaymentPlan)
            if repaymentFirst != 'nan':
                repayFirst = SubElement(loanterms, 'dac__repayment-first-date', iso_h_date=repaymentFirst)
            if repaymentFinal != 'nan':
                repayFinal = SubElement(loanterms, 'dac__repayment-final-date', iso_h_date=repaymentFinal)

        # dac:grant-equivalent value=["grant_equivalent"]
        try:
            grant = float(crsfile["grant_equivalent"][line])
        except ValueError:
            grant = 0
        grantAmount = str('{0:.2f}'.format(grant))
        if grantAmount != 0 and grantAmount != 'nan':
            grantEquivalent = SubElement(crsAdd, 'dac__grant-equivalent', value=grantAmount)

        # loan-status year="" value-date=""
        #   interest-received ["interest_received"]
        #   principal-outstanding ["Principa_disbursed"]
        #   principal-arrears ["Principal_arrears"]
        #   interest-arrears ["arrears_interest"]
        try:
            interestrec = float(crsfile["interest_received"][line])
            interestAmount = str('{0:.2f}'.format(interestrec))
        except ValueError:
            interestAmount = 'nan'
        try:
            prinout = float(crsfile["Principa_disbursed"][line])
            prinAmount = str('{0:.2f}'.format(prinout))
        except ValueError:
            prinAmount = 'nan'
        try:
            prinarr = float(crsfile["Principal_arrears"][line])
            prinarrAmount = str('{0:.2f}'.format(prinarr))
        except ValueError:
            prinarrAmount = 'nan'
        try:
            interestarr = float(crsfile["arrears_interest"][line])
            interestarrAmount = str('{0:.2f}'.format(interestarr))
        except ValueError:
            interestarrAmount = 'nan'
        reportingyear = ''
        if interestAmount != 'nan' or prinAmount != 'nan' or prinarrAmount != 'nan' or interestarrAmount != 'nan':
            if reportingyear != 'nan' and comdate != 'nan':
                loanstatus = SubElement(crsAdd, 'loan-status', value_h_date=comdate, year=str(reportyear))
            elif reportingyear != 'nan':
                loanstatus = SubElement(crsAdd, 'loan-status', year=str(reportyear))
            elif comdate != 'nan':
                loanstatus = SubElement(crsAdd, 'loan-status', value_h_date=comdate)
            else:
                loanstatus = SubElement(crsAdd, 'loan-status')
            if interestAmount != 'nan':
                intrec = SubElement(loanstatus, 'interest-received')
                intrec.text = interestAmount
            if prinAmount != 'nan':
                principaloutstanding = SubElement(loanstatus, 'principal-outstanding')
                principaloutstanding.text = prinAmount
            if prinarrAmount != 'nan':
                principalarrears = SubElement(loanstatus, 'principal-arrears')
                principalarrears.text = prinarrAmount
            if interestarrAmount != 'nan':
                interestarrears = SubElement(loanstatus, 'interest-arrears')
                interestarrears.text = interestarrAmount

        # channel-code ["channel_code"]
        try:
            channel = int(crsfile["channel_code"][line])
        except ValueError:
            channel = 0
        if channel != 0:
            channelcode = SubElement(crsAdd, 'channel-code')
            channelcode.text = str(channel)

        # dac:channel-description
        #   dac:narrative ["channel_name"]
        channelDesc = str(crsfile["channel_name"][line])
        if channelDesc != 'nan':
            channelDescription = SubElement(crsAdd, 'dac__channel-description')
            narrative = SubElement(channelDescription, 'narrative')
            narrative.text = channelDesc

        # dac:reporting-year ["reporting_year"]
        if reportyear != 0:
            reportingyear = SubElement(crsAdd, 'dac__reporting-year')
            reportingyear.text = str(reportyear)

        # dac:donorcode code=["reporting_country"]
        try:
            donorcode = int(crsfile["reporting_country"][line])
        except ValueError:
            donorcode = 0
        if donorcode != 0:
            reportingcountry = SubElement(crsAdd, 'dac__donorcode', code=str(donorcode))

        # dac:agency code=["extending_agency"]
        try:
            extending = int(crsfile["extending_agency"][line])
        except ValueError:
            extending = 0
        if extending != 0:
            extendingagency = SubElement(crsAdd, 'dac__agency', code=str(extending))

        # dac:nature-submission code=["nature_of_submission"]
        try:
            nature = int(crsfile["nature_of_submission"][line])
        except ValueError:
            nature = 0
        if nature != 0:
            naturesub = SubElement(crsAdd, 'dac__nature-submission', code=str(nature))

        # dac:commitment-date iso-date=["commitment_date"]
        if comdate != 0:
            commitdate = SubElement(crsAdd, 'dac__commitment-date', iso_h_date=comdate)

        # dac:currency code=["currency"]
        try:
            currency = int(crsfile["currency"][line])
        except ValueError:
            currency = 0
        if currency != 0:
            currencycode = SubElement(crsAdd, 'dac__currency', code=str(currency))

        # dac:other-amounts code=1 ["irtc"]
        #   dac:value value-date=""
        irtc = str(crsfile["irtc"][line])
        if irtc != 'nan':
            otheramount = SubElement(crsAdd, 'dac__other-amounts', code='1')
            irtcvalue = SubElement(otheramount, 'dac__value', value_h_date=comdate)
            irtcvalue.text = irtc

        # dac:other-amounts code=2 ["expert_commitment"]
        expertcom = ''
        expertcommit = str(crsfile["expert_commitment"][line])
        if expertcommit != 'nan':
            expertcom = SubElement(crsAdd, 'dac__other-amounts', code='2')
            expertvalue = SubElement(expertcom, 'dac__value', value_h_date=comdate)
            expertvalue.text = expertcommit

        # dac:other-amounts code=3 ["expert_extended"]
        expertextend = str(crsfile["expert_extended"][line])
        if expertextend != 'nan':
            expertext = SubElement(crsAdd, 'dac__other-amounts', code='3')
            expertextvalue = SubElement(expertext, 'dac__value', value_h_date=comdate)
            expertextvalue.text = expertextend

        # dac:other-amounts code=4 dac:value value-date="" ["export_credit"]
        export = str(crsfile["export_credit"][line])
        if export != 'nan':
            exportcred = SubElement(crsAdd, 'dac__other-amounts', code='4')
            exportvalue = SubElement(expertcom, 'dac__value', value_h_date=comdate)
            exportvalue.text = export

        # dac:mobilisation
        #   dac:mobilisation-leverage code=["Leverage_mech"]
        #   dac:mobilisation-origin code=["Orgin_of_funds"]
        #   dac:value ["Amounts_mobilized"]
        try:
            leverage = str(int(crsfile["Leverage_mech"][line]))
        except ValueError:
            leverage = 0
        try:
            origin = str(int(crsfile["Orgin_of_funds"][line]))
        except ValueError:
            origin = 0
        try:
            mobilvalue = float(crsfile["Amounts_mobilized"][line])
        except ValueError:
            mobilvalue = 0
        if leverage != 0 or origin != 0 or mobilvalue != 0:
            mobilisation = SubElement(crsAdd, 'dac__mobilisation')
            if leverage != 0:
                lev = SubElement(mobilisation, 'dac__mobilisation-leverage', code=leverage)
            if origin != 0:
                org = SubElement(mobilisation, 'dac__mobilisation-origin', code=origin)
            if mobilvalue != 'nan':
                mobvalue = SubElement(mobilisation, 'dac__value')
                # mobilAmount = '{0:.2f}'.format(mobilvalue)
                #mobvalue.text = mobilAmount
                mobvalue = mobilvalue

    # Make sure CRS-ADD has data in it. If not, just delete.
    for crsnode in activity.findall('crs-add'):
        node = crsnode.find('other-flags')
        if node is None:
            activity.remove(crsnode)

    for loannode in activity.findall('dac__loan-terms'):
        node = loannode.find('dac__repayment')
        print(node)
        if node is None:
            activity.remove(loannode)

# This is to write to a singular file.
if not os.path.exists('export/' + time.strftime("%m-%d-%Y") + '/'):
    os.makedirs('export/' + time.strftime("%m-%d-%Y") + '/')
output_file = open('export/' + time.strftime("%m-%d-%Y") + '/new_crs1.xml', 'w', encoding='utf-8')
output_file.write(prettify(activities).replace("__", ":").replace("_h_", "-"))
output_file.close()
print('Zipping...')
shutil.make_archive('export/zip/export-' + time.strftime("%m-%d-%Y"), 'zip', 'export/' +
                    time.strftime("%m-%d-%Y") + '/')
print('Complete!')
