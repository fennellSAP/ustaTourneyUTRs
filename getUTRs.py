from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import time
import operator

playersNotFound = []
driver = webdriver.Chrome(executable_path=r"C:\Users\I509049\Downloads\chromedriver_win32\chromedriver.exe")

def selectDivision():

    selectBox = driver.find_element_by_xpath('//*[@id="ctl00_mainContent_ControlTabs7_cboEvents"]')

    options = [x for x in selectBox.find_elements_by_tag_name("option")]

    print("\n\nEnter number corresponding to the division you'd like to select")

    counter = 1
    divisionDict = {}
    for element in options:

        if "-----" in element.text:
            continue

        divisionDict[counter] = element.text
        print(f'{counter}) {element.text}')
        counter += 1
    
    divisionSelect = int(input())
    return divisionDict[divisionSelect]

def getPlayers(division):

    print("Finding Players in USTA tournament...\n")

    select = Select(driver.find_element_by_xpath('//*[@id="ctl00_mainContent_ControlTabs7_cboEvents"]'))
    select.select_by_visible_text(division)

    time.sleep(1.25)

    totalPlayers = driver.find_element_by_xpath('//*[@id="ctl00_mainContent_ControlTabs7_pnlUpdate"]/div[8]').text
    totalPlayers = int(totalPlayers[15:])

    foundPlayers = 0
    rowNum = 3
    players = []

    while foundPlayers < totalPlayers:

        try:

            playerName = driver.find_element_by_xpath('//*[@id="applicants"]/tbody/tr[' + str(rowNum) + ']/td[1]/a[2]').text
            rowNum += 1
            foundPlayers += 1

            players.append(playerName)
    
        except:

            rowNum +=1 

    return players

def getUTRs(playersList):

    print("Finding tournament player's UTRs...\n")

    class TennisPlayer():
     def __init__(self):
         self.name = ""
         self.rating = 0
         self.location = ""

    playerData = []

    counter = 0
    for name in playersList:

        driver.find_element_by_xpath('//*[@id="myutr-app-wrapper"]/div[2]/nav/div[1]/div[2]/div/div[1]/div[1]/input').send_keys(name)
        
        try:
            element = WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located((By.XPATH, '//*[@id="myutr-app-wrapper"]/div[2]/nav/div[1]/div[2]/div/div[1]/div[2]/div/div/div[2]/div[2]/div[2]/span')))

            playerInfo = driver.find_element_by_xpath('//*[@id="myutr-app-wrapper"]/div[2]/nav/div[1]/div[2]/div/div[1]/div[2]/div/div/div[2]/div[2]/div[2]/span').text
            playerArray = playerInfo.split(u'\u2022')

            player = TennisPlayer()
            player.name = name

            if 'UR' in playerArray[1]:

                player.rating = 0

            else:
                
                player.rating = float(playerArray[1].strip())
                
            if len(playerArray) < 3:

                player.location = "Not Found"

            else:

                player.location = playerArray[2]

            playerData.append(player)
            
        except:
            
            player.location = "Not Found"
            playersNotFound.append(name)

        finally:

            driver.find_element_by_xpath('//*[@id="myutr-app-wrapper"]/div[2]/nav/div[1]/div[2]/div/div[1]/div[1]/input').send_keys(Keys.CONTROL + "a")
            driver.find_element_by_xpath('//*[@id="myutr-app-wrapper"]/div[2]/nav/div[1]/div[2]/div/div[1]/div[1]/input').send_keys(Keys.BACKSPACE)
            print(f'{counter}/{len(playersList)}')
            counter += 1

    playerData.sort(key=lambda x: x.rating, reverse=True)

    return playerData

def login():

    email = input("Enter UTR E-Mail: ")
    password = input("Enter UTR Password: ")

    print("Going to UTR website...")

    driver.get("https://www.myutr.com/")

    print("Logging in to UTR...\n")

    driver.find_element_by_xpath('//*[@id="myutr-app-wrapper"]/div[2]/nav/div/ul/li[4]/a').click()

    email = driver.find_element_by_xpath('//*[@id="emailInput"]').send_keys(email)
    password = driver.find_element_by_xpath('//*[@id="passwordInput"]').send_keys(password)

    driver.find_element_by_xpath('//*[@id="myutr-app-body"]/div/div/div/div/form/div[3]/button').click()

    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="myutr-app-wrapper"]/div[2]/nav/div[1]/div[4]/span')))
    
    except:
        return False

    finally:

        return True

        

def writeToExcel(playerData):

    print("Creating Excel File...\n")

    wb = Workbook()
    sheet = wb.active

    sheet['A1'] = "Seed"
    sheet['B1'] = "Name"
    sheet['C1'] = "Rating"
    sheet['D1'] = "Location"
    
    counter = 2

    for player in playerData:
        
        sheet["A" + str(counter)] = counter - 1
        sheet["B" + str(counter)] = player.name
        sheet["C" + str(counter)] = player.rating
        sheet["D" + str(counter)] = player.location

        counter += 1

    counter += 1
    sheet["A" + str(counter)] = "Players Who Returned No Data"
    counter += 2

    for nobody in playersNotFound:

        sheet["A" + str(counter)] = nobody
        counter += 1
    
    wb.save("playerSeedings.xlsx")


def main():

    tournamentURL = input("Enter tournament URL: ")

    tournamentURL += "#&&s=1"

    driver.get(tournamentURL)

    division = selectDivision()
    
    players = getPlayers(division)

    authenticated = login()

    if not authenticated:
        print("\nDid not pass authenticaton. Please try again")
        while not authenticated:
            authenticated = login()

    print("\nAuthentication Passed")    
    playerData = getUTRs(players)
    
    writeToExcel(playerData)

    
main()

