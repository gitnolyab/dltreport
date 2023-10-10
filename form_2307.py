import streamlit as st
from pathlib import Path
from docxtpl import DocxTemplate
import re
from datetime import datetime, timedelta
import time

hide_menu = """
                <style> #MainMenu { visibility:hidden; }
                           footer { visibility:hidden; }
                </style>
            """

# vperiodFrom = ''
# vperiodTo = ''
atc_desc = ''


# vbir_no = ''

# vMMfrom, vDDfrom, vYYYYfrom, vMMto, vDDto, vYYYYto,

# payeesName,yeesRegAddress


def process_2307(yeesZipCode, payeesName, yeesRegAddress, yorsTin,
                 payorsName, yorsRegAddress, yorsZipCode,
                 atc_code, firstQtr, secondQtr, thirdQtr,
                 vMMfrom, vDDfrom, vYYYYfrom, vMMto, vDDto,
                 vYYYYto, bir_no, corporation_name, vSignatory):
    # TO BE ADD UP payeesName, yeesRegAddress

    # OPEN FILE SOURCE FILE
    document_path = Path(__file__).parent / "FORM2307/TEMPLATE_2307.docx"
    doc = DocxTemplate(document_path)

    progress_bar = st.progress(0)

    # Simulate processing with sleep
    for i in range(1, 40):
        time.sleep(0.1)  # Simulate some work
        progress_bar.progress(i)

    # SOURCE CONTEXT:

    # FOR THE PERIOD:
    # DATE SPLIT - periodFrom
    # vperiodFrom = ''
    # vYYYYfrom, vMMfrom, vDDfrom = vperiodFrom.split('-')

    # DATE SPLIT - periodTo
    # vperiodTo = ''
    # vYYYYto, vMMto, vDDto = periodTo.split('-')

    # PART I - PAYEE's INFORMATION: 004 357 072,0000
    # bir_no
    vbir_no = bir_no.split('-')
    vPayeeTin1 = vbir_no[0]
    vPayeeTin2 = vbir_no[1]
    vPayeeTin3 = vbir_no[2]
    vPayeeTin4 = vbir_no[3]

    vPayeesName = payeesName
    vyeesRegAddress = yeesRegAddress
    vyeesZipCode = yeesZipCode

    # PART II - PAYOR's INFORMATION: 004 357 072,0000
    # ANG DATA KUHAON KANG PAYEE  yorsTin
    yorsTin = yorsTin.split('-')
    vPayorsTin1 = yorsTin[0]
    vPayorsTin2 = yorsTin[1]
    vPayorsTin3 = yorsTin[2]
    vPayorsTin4 = yorsTin[3]

    vPayorsName = payorsName
    vyorsRegAddress = yorsRegAddress
    vyorsZipCode = yorsZipCode

    vAtc_Code = atc_code
    vAtc_Desc = atc_desc

    # vFirstQtr = round(float(firstQtr), 2)
    # vFirstQtr = round(firstQtr, 2)
    # vSecondQtr = round(secondQtr, 2)
    # vThirdQtr = round(thirdQtr, 2)

    vFirstQtr = firstQtr
    vSecondQtr = secondQtr
    vThirdQtr = thirdQtr

    vTotal = 0.00
    vTWQ = 0.00
    atc_pct = 0.00

    if vAtc_Code == 'WI010':
        atc_pct = .1
    if vAtc_Code == 'WC010':
        atc_pct = .1
    if vAtc_Code == 'WI158':
        atc_pct = .01
    if vAtc_Code == 'WC158':
        atc_pct = .01
    if vAtc_Code == 'WI160':
        atc_pct = .02
    if vAtc_Code == 'WC160':
        atc_pct = .02
    if vAtc_Code == 'WI100':
        atc_pct = .05
    if vAtc_Code == 'WC100':
        atc_pct = .05
    if vAtc_Code == 'WI540':
        atc_pct = .05
    if vAtc_Code == 'WC540':
        atc_pct = .05
    if vAtc_Code == 'WI610':
        atc_pct = .01
    if vAtc_Code == 'WC610':
        atc_pct = .01

    print('atc_pct =', atc_pct)
    # Check which quarter has a value and calculate vTWQ accordingly
    if vFirstQtr != 0.0:
        vTotal = vFirstQtr
        vTWQ = vTotal * atc_pct
        print('vTWQ 1st', vTWQ)

    elif vSecondQtr != 0.0:
        vTotal = vSecondQtr
        vTWQ = vTotal * atc_pct
        print('vTWQ 2nd', vTWQ)

    elif vThirdQtr != 0.0:
        print('vThirdQtr ', vThirdQtr)
        vTotal = vThirdQtr
        vTWQ = vTotal * atc_pct
        print('vTWQ 3rd', vTWQ)

    # else:
    #   vTotal = 0.0

    # vTWQ = vTotal * atc_pct
    # vTotal = vFirstQtr + vSecondQtr + vThirdQtr
    # vTWQ = vTotal * atc_pct
    # vTWQ = round(float(vTotal), 2) * round(float(atc_pct), 2)
    # vTotal = round(float(vFirstQtr), 2) + round(float(vSecondQtr), 2) + round(float(vThirdQtr), 2)
    # vTWQ = round(vTotal * float(atc_pct), 2)

    # TARGET CONTEXT:
    target_context = {
        "FMM": vMMfrom,
        "FDD": vDDfrom,
        "FYYYY": vYYYYfrom,

        "TMM": vMMto,
        "TDD": vDDto,
        "TYYYY": vYYYYto,

        # DAPAT DIDTO NI MO PRINT SA PAYOR --------------

        # PART I - PAYOR'S INFORMATION NA DAPAT

        "PE01": vPayorsTin1,
        "PE02": vPayorsTin2,
        "PE03": vPayorsTin3,
        "PE04": vPayorsTin4,
        "PAYEES_NAME": vPayorsName,
        "PAYEES_REGISTERED_ADDRESS": vyorsRegAddress,
        "PEZIP": vyorsZipCode,

        # vPayeesName, vyeesRegAddress
        # Purok 4, Hubang, San Francisco, Philippines

        # END PRINT SA PAYOR UP --------------

        # PART II - PAYOR'S INFORMATION
        # OK NA PART II - PAYEES NA INFORMATION DAPAT

        "PR01": vPayeeTin1,
        "PR02": vPayeeTin2,
        "PR03": vPayeeTin3,
        "PR04": vPayeeTin4,

        "PAYORS_NAME": corporation_name,
        "PAYORS_REGISTERED_ADDRESS": "EDB BLDG 1, PRK 4, BRGY HUBANG, SAN FRANCISCO, AGUSAN DEL SUR",
        "PRZIP": '8501',
        # "PRZIP": vyeesZipCode

        "ATC_DESC": vAtc_Desc,
        "ATC_CODE": vAtc_Code,

        "FirstMoQ": "{:,.2f}".format(vFirstQtr),
        "SecondMoQ": "{:,.2f}".format(vSecondQtr),
        "ThirdMoQ": "{:,.2f}".format(vThirdQtr),
        "Total": "{:,.2f}".format(vTotal),
        "TWQ": "{:,.2f}".format(vTWQ),
        "SIGNATORY": vSignatory,

    }

    doc.render(target_context)
    # SAVE TO TARGET FILE
    doc.save(Path(__file__).parent / f"FORM2307/2307-OF={bir_no}.docx")

    # Update the progress bar to 100%
    progress_bar.progress(100)


def app():
    global atc_desc

    st.markdown(hide_menu, unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)

    # st.write("2307")
    # st.write("Certificate of Creditable Tax", )
    # st.write("Withheld at Source")
    # st.write('---')

    st.markdown("<div style='text-align: center; font-size: 36px;'>FORM 2307</div>", unsafe_allow_html=True)
    st.markdown("<div style='text-align: center; font-size: 24px;'>Certificate of Creditable Tax</div>",
                unsafe_allow_html=True)
    st.markdown("<div style='text-align: center; font-size: 24px;'>Withheld at Source</div>", unsafe_allow_html=True)

    st.write('---')
    # -----------------------

    st.write('FOR THE PERIOD')
    # Get the current date
    current_date = datetime.now()

    # Calculate the first day of the current month
    first_day_of_month = current_date.replace(day=1)

    # Calculate the last day of the current month
    next_month = current_date.replace(day=28) + timedelta(days=4)  # Move to next month and add 4 days
    last_day_of_month = next_month - timedelta(days=next_month.day)

    col1, col2 = st.columns(2)

    # Set the initial date values
    default_period_from = first_day_of_month.strftime("%Y-%m-%d")
    default_period_to = last_day_of_month.strftime("%Y-%m-%d")

    # Create the date inputs with the default values
    with col1:
        periodFrom = st.date_input("DATE FROM:", datetime.strptime(default_period_from, "%Y-%m-%d"),
                                   format="MM/DD/YYYY")
    with col2:
        periodTo = st.date_input("DATE TO:", datetime.strptime(default_period_to, "%Y-%m-%d"),
                                 format="MM/DD/YYYY")
    # -----------------------

    if periodFrom and periodTo:
        vYYYYfrom = periodFrom.year
        vMMfrom = str(periodFrom.month).zfill(2)  # Format to have leading zeros
        vDDfrom = str(periodFrom.day).zfill(2)  # Format to have leading zeros

        vYYYYto = periodTo.year
        vMMto = str(periodTo.month).zfill(2)  # Format to have leading zeros
        vDDto = str(periodTo.day).zfill(2)  # Format to have leading zeros

    st.write('')
    st.write('---')

    corporation_name = st.selectbox(
        "SELECT PAYOR'S NAME", ('',
                                'VASTMART CONVENIENCE STORE',
                                'DLT ROOFTOP RESTOBAR',
                                'JUNCTION 357 GASOLINE STATION',
                                'DELIS BAKESHOP',
                                'EUPHORIAS WELLNESS AND BEAUTY CENTER',
                                ))

    corp_bir_no = {
        'VASTMART CONVENIENCE STORE': "317-943-275-00001",
        'DLT ROOFTOP RESTOBAR': "317-943-275-000",
        'JUNCTION 357 GASOLINE STATION': "498-960-533-00000",
        'DELIS BAKESHOP': "480-443-905-00003",
        'EUPHORIAS WELLNESS AND BEAUTY CENTER': "480-443-905-00000",
    }
    if corporation_name in corp_bir_no:  # e search niya sa  dictionary
        bir_no = corp_bir_no[corporation_name]  # kung makita replace with context
        st.write(bir_no)

    st.write('---')

    payeesName = ''
    yeesRegAddress = ''
    yeesZipCode = ''

    st.write("PAYEE's INFORMATION:")
    payorsName = st.text_input("Payee's Name: ", placeholder="enter payee's name here...")
    yorsRegAddress = st.text_input("Payee's Registered Address: ", placeholder="enter payee's address here...")

    col1, col2 = st.columns(2)

    with col1:
        yorsTin = st.text_input("Payee's Tin Number: format (000-000-000-0000)", placeholder="enter payee's tin number")
    # Define a regex pattern to match the expected format
    tin_pattern = r'^\d{3}-\d{3}-\d{3}-[0-9]*0\d*$'
    # Check if the TIN input is valid
    if not re.match(tin_pattern, yorsTin) and yorsTin:
        st.error("Invalid TIN format. Please use the format 000-000-000-0000.")
    else:
        # Clear any previous error messages
        st.empty()

    with col2:
        yorsZipCode = st.text_input("Payee's Zip Code: ", placeholder="enter payee's zip code here...")

    # Define a regex pattern to match numeric input
    numeric_pattern = r'^\d+$'
    # Check if the Zip Code input is valid (numeric)
    if not re.match(numeric_pattern, yorsZipCode) and yorsZipCode:
        st.error("Invalid Zip Code format. Please enter a numeric value.")
    else:
        # Clear any previous error messages
        st.empty()

    st.write('---')
    st.write('')

    st.write("DETAILS OF MONTHLY INCOME AND TAXES WITHHELD")
    st.write("Income Payment Subject to Expanded Withholding Tax")
    st.write('')
    st.write('')
    atc_code = st.selectbox(
        'ATC CODE', ('', 'WI010', 'WC010',
                     'WI158', 'WC158',
                     'WI160', 'WC160',
                     'WI100', 'WC100',
                     'WI540', 'WC540',
                     'WI610', 'WC610',))

    atc_code_mapping = {
        'WI010': "Professional (Lawyers, CPA's, Engineers,etc) 10.00%",
        'WC010': "Professional (Lawyers, CPA's, Engineers,etc) 10.00%",
        'WI158': "Income payment made by top withholding agents to their local/resident supplier of goods other than \
                   those covered by other rates of withholding tax	1.00%",
        'WC158': "Income payment made by top withholding agents to their local/resident supplier of goods other than \
                   those covered by other rates of withholding tax	1.00%",
        'WI160': "Income payment made bytop withholding agents to their local resident \
                   supplier of services other than those covered by other rates of \
                    withholding tax	2.00% ",
        'WC160': "Income payment made bytop withholding agents to their local resident \
                   supplier of services other than those covered by other rates of \
                    withholding tax	2.00%",
        'WI100': "Rentals: On gross rental or lease for the continued use or possession of personal property \
                   in excess of Ten Thousand Pesos (P10,000.00) annually and real property used in business \
                   which the payor or obligator has not taken title or is not taking title, or in which has \
                   no equity; poles, satellites, transmission facilities and billboards)  5.00% ",
        'WC100': "Rentals: On gross rental or lease for the continued use or possession of personal property \
                   in excess of Ten Thousand Pesos (P10,000.00) annually and real property used in business \
                   which the payor or obligator has not taken title or is not taking title, or in which has \
                   no equity; poles, satellites, transmission facilities and billboards)  5.00% ",
        'WI540': "Tolling fees paid to refineries-corporate	5.00%",
        'WC540': "Tolling fees paid to refineries-corporate	5.00%",
        'WI610': "Income payments made to suppliers of agricultural products in excess of cumulative \
                    amount of P300,000.00 within the same taxable year	1.00% ",
        'WC610': "Income payments made to suppliers of agricultural products in excess of cumulative \
                    amount of P300,000.00 within the same taxable year	1.00% ",
    }
    if atc_code in atc_code_mapping:  # e search niya sa month_mapping dictionary
        atc_desc = atc_code_mapping[atc_code]  # kung makita replace with context
        st.write(atc_desc)

    col1, col2, col3 = st.columns(3)
    with col1:
        firstQtr = st.number_input('1st Month of the Quarter')
    with col2:
        secondQtr = st.number_input('2nd  Month of the Quarter')
    with col3:
        thirdQtr = st.number_input('3rd Month of the Quarter')

    st.write('---')

    # default_signatory = "DESERIE AMON"
    # signatory = st.text_input("Signatory:", value=default_signatory)

    signatory_name = st.selectbox(
        "SELECT YOUR NAME --- >>> FOR SIGNATORY", ('',
                                                   'REGINE GECARANE',
                                                   'OTHELIA MAE A. CASTULO',
                                                   'DESERIE B. AMON',
                                                   'JESSIEL L. CALAWIGAN',
                                                   'JENNY SHIPLEY B. ANGCLA',
                                                   ))

    # signatory = st.text_input("Signatory: ", placeholder='Please enter Signatory Name')
    vSignatory = signatory_name

    st.write('---')

    button_clicked = st.button('PRINT FORM 2307')
    # Validation

    try:
        if button_clicked:
            process_2307(yeesZipCode, payeesName, yeesRegAddress, yorsTin,
                         payorsName, yorsRegAddress, yorsZipCode,
                         atc_code, firstQtr, secondQtr, thirdQtr,
                         vMMfrom, vDDfrom, vYYYYfrom, vMMto, vDDto,
                         vYYYYto, bir_no, corporation_name, vSignatory)

            st.success(f'{corporation_name} {bir_no} CREATED ----- >>> Please check the file...')
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        st.success('please complete the entry...')

    # TO BE ADDED payeesName, yeesRegAddress HERE UP

    # OR YOU CAN PRINT 2307 AUTOMATION THRU EXCEL FILE
    st.write('---')
    # ----------------------------------------------------


'''
# print ms word directly
    
    import subprocess
    
    # Define the path to the Word document you want to print
    document_path = "D:/sourcefile/your_document.docx"
    
    # Define the Microsoft Word executable path (replace with your actual path)
    msword_path = "C:/Program Files/Microsoft Office/Office16/WINWORD.EXE"
    
    # Define the print command
    print_command = f'"{msword_path}" /t "{document_path}"'
    
    # Execute the print command
    try:
        subprocess.run(print_command, shell=True)
        print("Printing command sent successfully.")
    except Exception as e:
        print(f"Error sending printing command: {str(e)}")

'''
