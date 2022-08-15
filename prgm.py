from urllib import request, error
import os
import zipfile
import pandas
import pprint

playerdata = {}

year = 2022 #The current year.
data_incomplete = True #Set to false if we have all the excel files we need and they're up to date

if data_incomplete == True:
    for i in range(2000, year): #Downloads and unzips all pre-current-year excel files if they don't exist 
        if not (os.path.exists(str(i) + ".xls") or os.path.exists(str(i) + ".xlsx")):
            try:
                request.urlretrieve("http://www.tennis-data.co.uk/{}/{}.zip".format(i, i), "{}.zip".format(i))
                zipped = zipfile.ZipFile("{}.zip".format(i), "r")
                zipped.extractall()
                zipped.close()
                os.remove("{}.zip".format(i))
                print("Downloaded and unzipped data for {}.".format(i))
            except error.HTTPError:
                request.urlretrieve("http://www.tennis-data.co.uk/{}/{}.xlsx".format(i, i), "{}.xlsx".format(i))
                print("Downloaded data for {}.".format(i))
        else:
            print("Data for {} already exists.".format(i))
    if os.path.exists(str(year) + ".xlsx"): #Deletes the excel file for the current year if it exists.
        os.remove("{}.xlsx".format(year))
    elif os.path.exists(str(year) + ".xls"):
        os.remove("{}.xls".format(year))
    request.urlretrieve("http://www.tennis-data.co.uk/{}/{}.xlsx".format(year, year), "{}.xlsx".format(year)) #Downloads and unzips (updates) current year excel file.
    print("Downloaded and unzipped current data for {}.".format(year))
    print("Data collection finished.")
else:
    print("Skipping data collection.")

data = {}
for yr in range(2000, year + 1):
    print("Converting excel data for {}".format(yr))
    try:
        currentfile = pandas.ExcelFile("{}.xls".format(yr))
    except:
        currentfile = pandas.ExcelFile("{}.xlsx".format(yr))
    print("Parsing excel data for {}".format(yr))
    data[yr] = currentfile.parse(currentfile.sheet_names[0]) #Parses first sheet of the excel file and saves to dictionary.
finaldata = {}
for yr in data:
    playerdata[yr] = {}
    for index in range(data[yr].shape[0]):
        winner = data[yr].iloc[index]["Winner"].strip().lower()
        loser = data[yr].iloc[index]["Loser"].strip().lower()
        if winner not in playerdata[yr]:
            playerdata[yr][winner] = 1000
        if loser not in playerdata[yr]:
            playerdata[yr][loser] = 1000
        k = 100
        playerdata[yr][winner], playerdata[yr][loser] = round(playerdata[yr][winner] + k * (1 - (1 / (1 + pow(10, (playerdata[yr][loser] - playerdata[yr][winner]) / 400)))), 2), round(playerdata[yr][loser] + k * (0 - (1 / (1 + pow(10, (playerdata[yr][winner] - playerdata[yr][loser]) / 400)))), 2)
    scoreslist = playerdata[yr].items()
    newscoreslist = []
    for score in scoreslist:
        print((score, yr))
        if len(newscoreslist) == 0:
            newscoreslist.append(score)
        else:
            if score[1] > newscoreslist[0][1]:
                newscoreslist.insert(0, score)
            else:
                if score[1] <= newscoreslist[-1][1]:
                    newscoreslist.append(score)
                else:
                    for index in range(len(newscoreslist)):
                        if newscoreslist[index][1] <= score[1]:
                            newscoreslist.insert(index, score)
                            break
    finaldata[yr] = newscoreslist

pprint.pprint(finaldata)
writefile = open("log.txt", "w")
pprint.pprint(finaldata, stream=writefile)
writefile.close()
print("Written to text file.")

#CSV file
print("Generating CSV file.")
csv = open("tennisdata.csv", "w")
csv.write("Player,Year,EloRating\n")
for year in finaldata:
    for name, score in finaldata[year]:
        csv.write("{},{},{}\n".format(name, year, score))
csv.close()



    
    

