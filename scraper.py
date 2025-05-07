import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Util to input "N/A" in the cell if the value is blank
def check(value):
  if value == None:
    return 'N/A'
  else:
    return value.getText()

# Util to format SHL and BHL location cells depending on what data is present
def lots(str, lot, footage_ns, footage_ew):
  if str == None:
    return "N/A"
  elif str != None and lot != None:
    return f"{str.getText()}  {lot.getText()}  {footage_ns.getText()}  {footage_ew.getText()}"
  else:
    return f"{str.getText()}  {footage_ns.getText()}  {footage_ew.getText()}"

# Source csv
csv = pd.read_csv('Wells_Locations.csv')

# Set up headless Chrome
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")  # Optional: for Windows systems
chrome_options.add_argument("--no-sandbox") 

# Set up web driver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Excel output column headers
headers = [
  'API Number',
  'Current Operator',
  'Well Name',
  'Well Number',
  'Well Type',
  'Well Direction',
  'Well Status',
  'Section',
  'Township',
  'Range',
  'OCD Unit Letter',
  'Surface Location',
  'Bottomhole Location',
  'Formation',
  'MD',
  'TVD',
  'Operator Name',
  'Operator Address',
  'Operator City, State, ZIP Code',
  'Well Details',
  'Well Files',
  'Tracking Number',
  'Shipper'
]

# List to store scraped well information
details = []
  
rows_completed = []

for index, row in csv.iterrows():
  '''This is new naming'''
  well_details = row['Link to Well Details']
  well_files = row['Link to Scanned Well Files']

  while index not in rows_completed:
    try:
      driver.get(well_details)
      wait = WebDriverWait(driver, 5)
      wait.until(EC.presence_of_element_located((By.ID, "ctl00_ctl00__main_main_lblApi"))) 
      data = BeautifulSoup(driver.page_source, features="html.parser")
      
      # Well Name and Number
      well_name_and_number = data.find(id="ctl00_ctl00__main_main_lblApi").getText()
      well_name_and_number = well_name_and_number.strip().rsplit(" ", 1)

      # API Number, Well Name, and Well Number
      api_number = well_name_and_number[0].strip().split(" ", 1)[0]
      well_name = well_name_and_number[0].strip().split(" ", 1)[1]
      well_number = well_name_and_number[1]

      # Current Operator
      operator_name = check(data.find(id="ctl00_ctl00__main_main_ucOperator_lblOgridName"))
      operator_name_a = check(data.find(id="ctl00_ctl00__main_main_ucGeneralWellInformation_lblOperator"))

      # Well Type
      well_type = check(data.find(id="ctl00_ctl00__main_main_ucGeneralWellInformation_lblWellType"))

      # Well Direction
      well_direction = check(data.find(id="ctl00_ctl00__main_main_ucGeneralWellInformation_lblDirectionalStatus"))

      # Well Status
      well_status = check(data.find(id="ctl00_ctl00__main_main_ucGeneralWellInformation_lblStatus"))

      # Section, Township, Range
      section_township_range = data.find(id="ctl00_ctl00__main_main_ucGeneralWellInformation_Location_lblLocation").getText().split("-")
      [ OCDUnitLetter, Section, Township, Range ] = section_township_range 
      ocd_unit_letter = OCDUnitLetter
      section = Section
      township = f"T{Township}"
      rng = f"R{Range}" 

      # Surface Location
      surface_hole_location = lots(data.find(id="ctl00_ctl00__main_main_ucGeneralWellInformation_Location_lblLocation"), data.find(id="ctl00_ctl00__main_main_ucGeneralWellInformation_Location_lblLot"), data.find(id="ctl00_ctl00__main_main_ucGeneralWellInformation_Location_lblFootageNSH"), data.find(id="ctl00_ctl00__main_main_ucGeneralWellInformation_Location_lblFootageEW"))

      # Bottomhole Location
      bottom_hole_location = lots(data.find(id="ctl00_ctl00__main_main_ucWellCompletions_rptCompletionsSummary_ctl00_Location_lblLocation"), data.find(id="ctl00_ctl00__main_main_ucWellCompletions_rptCompletionsSummary_ctl00_Location_lblLot"), data.find(id="ctl00_ctl00__main_main_ucWellCompletions_rptCompletionsSummary_ctl00_Location_lblFootageNSH"), data.find(id="ctl00_ctl00__main_main_ucWellCompletions_rptCompletionsSummary_ctl00_Location_lblFootageEW"))
  
      # Formation
      formation = "%s %s" %(
        check(data.find(id="ctl00_ctl00__main_main_ucWellCompletions_rptCompletionsSummary_ctl00_lblWellInformationPool")),
        check(data.find(id="ctl00_ctl00__main_main_ucWellCompletions_rptCompletionsSummary_ctl00_lblWellInformationPoolName"))
      )

      # MD
      measured_depth = check(data.find(id="ctl00_ctl00__main_main_ucGeneralWellInformation_lblMeasuredVerticalDepth"))

      # TVD
      total_vertical_depth = check(data.find(id="ctl00_ctl00__main_main_ucGeneralWellInformation_lblTrueVerticalDepth"))#

      # Operator Name
      operator_name = check(data.find(id="ctl00_ctl00__main_main_ucOperator_lblOgridName"))

      # Operator Address
      operator_address = check(data.find(id="ctl00_ctl00__main_main_ucOperator_lblAddress"))
      
      # Operator City
      operator_city = check(data.find(id="ctl00_ctl00__main_main_ucOperator_lblCityStateZip"))

      # Tracking Number
      tracking_number = ""

      # Shipper
      shipper = ""

      details.append([
        api_number,
        operator_name_a,
        well_name,
        well_number,
        well_type,
        well_direction,
        well_status,
        section,
        township,
        rng,
        ocd_unit_letter,
        surface_hole_location,
        bottom_hole_location,
        formation,
        measured_depth,
        total_vertical_depth,
        operator_name,
        operator_address,
        operator_city,
        well_details,
        well_files,
        tracking_number,
        shipper
      ])
      
      rows_completed.append(index)
      print("{:.0f}% completed".format(((index + 1) / float(len(list(csv.iterrows())))) * 100))
    except Exception as e:
      pass
    
driver.quit()
filename = input('Pick a file name: ')
output = pd.DataFrame(details, columns=headers)
output.set_index('API Number', inplace=True)
output.to_excel(f'{filename}.xlsx', sheet_name="Wells")


