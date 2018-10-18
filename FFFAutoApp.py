y = input('Enter a number of records to process: ')
rn = input('Enter row number to start from: ')
x = 1
rn1 = int(rn) 
while x <= int(y): #specify number of times to run process
    import os
    from datetime import datetime, date
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support.ui import Select
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import NoSuchElementException
    from selenium.common.exceptions import ElementNotInteractableException
    from openpyxl import Workbook
    from openpyxl import load_workbook
    wb1 = load_workbook('FFFInputData.xlsx')
    ws1 = wb1['AppDataList']
    error = (NoSuchElementException, ElementNotInteractableException)
    browser = webdriver.Firefox(executable_path=r'C:\Users\garren-james\AppData\Local\geckodriver')
    browser.get('http://foodforflorida-acc.dcf.state.fl.us/foodforflorida/FrmWorkerLogin.do?performAction=init')
    col1 = 'C'+str(rn1)     #PIP First Name
    col2 = 'D'+str(rn1)     #PIP Last Name
    col3 = 'E'+str(rn1)     #PIP User ID
    col4 = 'F'+str(rn1)     #PIP SSN
    col5 = 'G'+str(rn1)     #PIP DOB
    col6 = 'H'+str(rn1)     #PIP Gender
    col7 = 'I'+str(rn1)    #Child First Name
    col8 = 'J'+str(rn1)    #Child DOB
    col9 = 'K'+str(rn1)    #Child Gender
    col10 = 'L'+str(rn1)    #City
    col11 = 'M'+str(rn1)    #zip
    col12 = 'N'+str(rn1)    #PIP Race
    col13 = 'O'+str(rn1)   #Child Race
    col14 = 'P'+str(rn1)    #Has Income?
    col15 = 'Q'+str(rn1)    #Income Amount
    col16 = 'R'+str(rn1)    #Pay Freq
    col17 = 'S'+str(rn1)    #Number of Hours
    col18 = 'T'+str(rn1)    #Start Date
    col19 = 'U'+str(rn1)   #Unearned Income?
    col20 = 'V'+str(rn1)   #Unearned Amount
    col21 = 'W'+str(rn1)   #Unearned Start Date
    col26 = 'X'+str(rn1)   ##Current DL #
    col27 = 'Y'+str(rn1)   ##1ST 3 SSN
    col28 = 'Z'+str(rn1)   ##2ND 2 SSN
    col29 = 'AA'+str(rn1)   ##LAST 3 ssn
    col30 = 'AB'+str(rn1)   ##Month DOB
    col31 = 'AC'+str(rn1)   ##Day DOB
    col32 = 'AD'+str(rn1)   ##Year DOB
    col33 = 'AE'+str(rn1)   ##County
    col34 = 'B'+str(rn1)    ##App Num
    col35 = 'AG'+str(rn1)   ##Last 4 of Child SSN
    col36 = 'AH'+str(rn1)
    col37 = 'AI'+str(rn1)
    col38 = 'AJ'+str(rn1)
    col39 = 'AK'+str(rn1)
    col40 = 'AL'+str(rn1)
    col41 = 'AM'+str(rn1)
    ## Following variables are workbook/sheet cell variables
    fn1 = ws1[col1]     #PIP First Name
    ln1 = ws1[col2]     #PIP Last Name
    ui1 = ws1[col3]     #PIP User ID
    sn1 = ws1[col4]     #PIP SSN
    db1 = ws1[col5]     #PIP DOB
    gd1 = ws1[col6]     #PIP Gender
    cfn1 = ws1[col7]    #Child First Name
    cdb1 = ws1[col8]    #Child DOB
    cgd1 = ws1[col9]    #Child Gender
    ct1 = ws1[col10]    #City
    zp1 = ws1[col11]    #zip
    rc1 = ws1[col12]    #PIP Race
    crc1 = ws1[col13]   #Child Race
    hi1 = ws1[col14]    #Has Income?
    ia1 = ws1[col15]    #Income Amount
    fq1 = ws1[col16]    #Pay Freq
    nh1 = ws1[col17]    #Number of Hours
    id1 = ws1[col18]    #Start Date
    uni1 = ws1[col19]   #Unearned Income?
    una1 = ws1[col20]   #Unearned Amount
    und1 = ws1[col21]   #Unearned Start Date
    dl1 = ws1[col26]    #DL Info 
    snp1 = ws1[col27]   #ssn 1st 3
    snp2 = ws1[col28]   #ssn next 2
    snp3 = ws1[col29]   #ssn last 4
    dbp1 = ws1[col30]   #Month of birth
    dbp2 = ws1[col31]   #Day of birth
    dbp3 = ws1[col32]   #Year of birth
    cty1 = ws1[col33]   #County Name
    csn4 = ws1[col35]   #Child SSN
    cd1= ws1[col36]
    cd2 = ws1[col37]
    cd3 = ws1[col38]
    idsel1 = ws1[col39]     #Has FLA ID Choice
    idsel2 = ws1[col40]     #If Not FLA, type
    idsel3 = ws1[col41]     #When not FLA, Other ID #

    ## Following variables are specific cell values 
    fnv1 = fn1.value
    lnv1 = ln1.value
    uiv1 = ui1.value
    snv1 = sn1.value
    if db1.value is None:
        pass
    else:
        dbv1 = db1.value.strftime("%m%d%Y")


    if gd1.value is None:
        gdv1 = input('Enter the Gender of {}: '.format(fnv1)).upper()
    else:
        gdv1 = gd1.value.upper()


    cfnv1 = cfn1.value
    if cdb1.value is None:
        pass
    else:
        cdbv1 = cdb1.value.strftime("%m%d%Y")


    if cgd1.value is None:
        cgdv1 = input('Enter the Gender of {}: '.format(cfnv1)).upper()
    else:
        cgdv1 = cgd1.value.upper()



    ctv1 = ct1.value
    zpv1 = zp1.value
    if rc1.value is None:
        pass
    else:
        rcv1 = rc1.value.upper()


    if crc1.value is None:
        pass
    else:
        crcv1 = crc1.value.upper()


    if hi1.value is None:
        pass
    else:
        hiv1 = hi1.value.upper()


    iav1 = ia1.value
    if fq1.value is None:
        fqv1 = fq1.value
    else:
        fqv1 = fq1.value.upper()


    nhv1 = nh1.value
    if id1.value is None or id1.value == "":
        pass
    else:
        idv1 = id1.value.strftime("%m%d%Y")


    if uni1.value is None:
        pass
    else:
        univ1 = uni1.value.upper()


    unav1 = una1.value
    if und1.value is None or und1.value == "":
        pass
    else:
        undv1 = und1.value.strftime("%m%d%Y")


    dlv1 = dl1.value
    snpv1 = snp1.value
    snpv2 = snp2.value
    snpv3 = snp3.value
    dbpv1 = dbp1.value
    dbpv2 = dbp2.value
    dbpv3 = dbp3.value
    ctyv1 = cty1.value
    csnv4 = csn4.value
    cdv1 = cd1.value
    cdv2 = cd2.value
    cdv3 = cd3.value
    idselv1 = idsel1.value
    if idsel2.value is None or idsel2.value == "":
        pass
    else:
        idselv2 = idsel2.value


    if idsel3.value is None or idsel3.value == "":
        pass
    else:
        idselv3 = str(idsel3.value)


    htmlElem = browser.find_element_by_id('selectprereg')
    htmlElem.click()
    htmlElem = browser.find_element_by_name('btnContinue')
    htmlElem.click()
    htmlElem = browser.find_element_by_id('applicantfname')
    htmlElem.send_keys(fnv1)
    htmlElem = browser.find_element_by_id('applicantlname')
    htmlElem.send_keys(lnv1)
    htmlElem = browser.find_element_by_id('applicantssn1')
    htmlElem.send_keys(snpv1)
    htmlElem = browser.find_element_by_name('applicantssn2')
    htmlElem.send_keys(snpv2)
    htmlElem = browser.find_element_by_name('applicantssn3')
    htmlElem.send_keys(snpv3)
    htmlElem = browser.find_element_by_id('applicantdobmm')
    htmlElem.send_keys(dbpv1)
    htmlElem = browser.find_element_by_name('applicantdobdd')
    htmlElem.send_keys(dbpv2)
    htmlElem = browser.find_element_by_name('applicantdobccyy')
    htmlElem.send_keys(dbpv3)
    htmlElem = browser.find_element_by_name('applicantsex')
    if gdv1 == 'M':
        htmlElem.send_keys(Keys.DOWN)


    if gdv1 == 'F':
        htmlElem.send_keys(Keys.DOWN)
        htmlElem.send_keys(Keys.DOWN)


    htmlElem = browser.find_elements_by_name('idtypeFlag')
    if idselv1 == 'Y':
        htmlElem = htmlElem[0].click()
        htmlElem = browser.find_element_by_id('applicantID')
        htmlElem.send_keys(dlv1)
        htmlElem = browser.find_element_by_name('btnContinue')
        htmlElem.click()
    else:
        htmlElem = htmlElem[1].click()
        htmlElem = browser.find_element_by_name('btnContinue')
        htmlElem.click()
        htmlElem = browser.find_element_by_name('btnContinue')
        htmlElem.click()
        htmlElem = browser.find_element_by_id('complsecque')
        if idselv2 == 'Out of State':
            htmlElem.send_keys(Keys.DOWN)


        if idselv2 == 'Passport':
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)


        if idselv2 == 'INS Doc':
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)


        if idselv2 == 'School/Work ID':
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)


        if idselv2 == 'USCIS Card':
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)


        if idselv2 == 'Other':
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)
            htmlElem.send_keys(Keys.DOWN)


        htmlElem = browser.find_element_by_name('idNumber')
        htmlElem.send_keys(idselv3)
        htmlElem = browser.find_element_by_name('btnContinue')
        htmlElem.click()


    table_element = browser.find_elements_by_tag_name('b')[-1]
    cellloc = table_element.get_attribute('innerHTML')
    cellloc = int(cellloc)
    ws1[col34] = cellloc
    wb1.save('FFFInputData.xlsx')
    htmlElem = browser.find_element_by_name('password1')
    htmlElem.send_keys('test123')
    htmlElem = browser.find_element_by_name('password2')
    htmlElem.send_keys('test123')
    htmlElem = browser.find_element_by_name('applsecurityqtn')
    htmlElem.send_keys(Keys.DOWN)
    htmlElem = browser.find_element_by_name('applsecurityans')
    htmlElem.send_keys('aa')
    htmlElem = browser.find_element_by_name('btnContinue')
    htmlElem.click()
    htmlElem = browser.find_element_by_id('disasterStreetNumber')
    htmlElem.send_keys('1000')
    htmlElem = browser.find_element_by_id('disasterRuralBoxStreetName')
    htmlElem.send_keys('anywhere')
    htmlElem = browser.find_element_by_id('disasterStreetSuffixText')
    htmlElem.send_keys(Keys.DOWN)
    htmlElem = browser.find_element_by_id('disasterCityName')
    htmlElem.send_keys(ctv1)
    htmlElem = browser.find_element_by_id('disasterZip5Code')
    htmlElem.send_keys(zpv1)
    htmlElem = Select(browser.find_element_by_id('disasterCountyCode'))
    htmlElem.select_by_value(ctyv1)
    htmlElem = browser.find_element_by_id('disasterid')
    htmlElem.send_keys(Keys.DOWN)
    htmlElem.send_keys(Keys.DOWN)
    htmlElem.send_keys(Keys.DOWN)
    htmlElem = browser.find_elements_by_name('applaltaddressflag')[1]
    htmlElem.click()
    htmlElem = browser.find_elements_by_name('specialNeeds')[1]
    htmlElem.click();
    htmlElem = browser.find_element_by_name('disasterdamagesw')
    htmlElem.click()
    htmlElem = browser.find_element_by_name('btnContinue')
    htmlElem.click()
    htmlElem = browser.find_elements_by_tag_name('button')[1]
    htmlElem.click()
    htmlElem = browser.find_elements_by_id('indvfname')[-1]
    htmlElem.send_keys(cfnv1)
    htmlElem = browser.find_elements_by_id('indvlname')[-1]
    htmlElem.send_keys(lnv1)
    htmlElem = browser.find_elements_by_id('indvssn1')[-1]
    htmlElem.send_keys(snpv1)
    htmlElem = browser.find_element_by_name('individualsId[1].indvssn2')
    htmlElem.send_keys(snpv2)
    htmlElem = browser.find_element_by_name('individualsId[1].indvssn3')
    htmlElem.send_keys(csnv4)
    htmlElem = browser.find_elements_by_id('indvdobmm')[-1]
    htmlElem.send_keys(cdv1)
    htmlElem = browser.find_element_by_name('individualsId[1].indvdobdd')
    htmlElem.send_keys(cdv2)
    htmlElem = browser.find_element_by_name('individualsId[1].indvdobccyy')
    htmlElem.send_keys(cdv3)
    htmlElem = browser.find_elements_by_id('indvsex')[-1]
    if cgdv1 == 'M':
        htmlElem.send_keys(Keys.DOWN)


    if cgdv1 == 'F':
        htmlElem.send_keys(Keys.DOWN)
        htmlElem.send_keys(Keys.DOWN)


    htmlElem = browser.find_element_by_name('btnContinue')
    htmlElem.click()
    htmlElem = browser.find_element_by_id('assettype')
    htmlElem.send_keys(Keys.DOWN)
    htmlElem = browser.find_element_by_id('assetamount')
    htmlElem.send_keys('250.00')
    htmlElem = browser.find_element_by_id('expensetype')
    htmlElem.send_keys(Keys.DOWN)
    htmlElem = browser.find_element_by_id('expenseamount')
    htmlElem.send_keys('300.00')
    htmlElem = browser.find_elements_by_tag_name('button')[2]
    htmlElem.click()
    htmlElem = browser.find_elements_by_id('expensetype')[-1]
    htmlElem.send_keys(Keys.DOWN)
    htmlElem.send_keys(Keys.DOWN)
    htmlElem.send_keys(Keys.DOWN)
    htmlElem.send_keys(Keys.DOWN)
    htmlElem = browser.find_elements_by_id('expenseamount')[-1]
    htmlElem.send_keys('100.00')
    if iav1 == None:
        pass
    else:
        htmlElem = browser.find_element_by_id('incomename')
        htmlElem.send_keys(Keys.DOWN)
        htmlElem = browser.find_element_by_id('incometype')
        htmlElem.send_keys(Keys.DOWN)
        htmlElem = browser.find_element_by_id('incomeamount')
        htmlElem.send_keys(iav1)


    htmlElem = browser.find_element_by_name('disasterpurchfood')
    htmlElem.click()
    #htmlElem = browser.find_element_by_name('btnContinue')
    #htmlElem.click()
    #htmlElem = browser.find_element_by_name('btnQuit')
    #htmlElem.click()
    #alert = browser.switch_to.alert
    #alert.accept()
    htmlElem = browser.find_element_by_name('btnContinue')
    htmlElem.click()
    htmlElem = browser.find_element_by_name('btnContinue')
    htmlElem.click()
    htmlElem = browser.find_element_by_name('btnNext')
    htmlElem.click()
    htmlElem = browser.find_element_by_name('btnExit')
    htmlElem.click()
    browser.quit()
    rn1 = rn1+1
    x = x+1
    print('Next Record: ' + str(x))
