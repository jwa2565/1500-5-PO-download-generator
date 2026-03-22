#version notes
#BETA v2.2: Made the character field 60 characters long, added today's date to the filenames, and adjusted the fab macro calculations. 
#           Oh, and added a delay to the end of the program so peeps can read the text that appears in the terminal lol
#           Also, note for later. Marco wants qtys added. 
#           Also also, adjust the label file to print 2 labels when using the 4 rail profileIDs

#BETA v3.0: Added new field to PO Download Spreadsheet. Thie new field is to specify an exact scannable order number for the scannable barcode for the end of line stuff

#BETA v3.2: Added balance info to the spreadsheet to adjust for BALANCEHOLE macro. Also made sure all macro logics (except tapcon) were up to date

#BETA v3.3: (potential) QUality of life improvements:
#   -Use the Schedule date in the filename. DONE
#   -Maybe add the filecounter to the filename?? DONEEEEEEEEEE
#   -Fix glitch when OrderNumber or Full SCannable order number are just numbers and not a string DONEEEEE
#   -CAN I PLEASE FIX HAVING TO REMOVE HALF OF THE SASHES :'(  ----OK KINDA DONE..... it got messy af AND most likely wont work for sliders. 

print("Starting 1500-5 PO Download Generator!")
print("Importing data...")

import re
import time
from datetime import date
today = str(date.today())
#print(today)
import pandas as pd
import os

df = pd.read_excel("PO Download Spreadsheet.xlsx")

print("Data import successful!")
print("Generating PO download files...")


        
def cleanDataframe():
    
    df["Order Number"] = df["Order Number"].fillna("empty")
    df["Customer"] = df["Customer"].fillna("###")
    df["Schedule Date"] = df["Schedule Date"].fillna("###")
    df["Destination"] = df["Destination"].fillna("###")
    df["Mullion"] = df["Mullion"].fillna("Single") #making Single the default option
    df["Child Safety Latch"] = df["Child Safety Latch"].fillna("No") #making No the default option
    df["Full Scannable Order Number"] = df["Full Scannable Order Number"].fillna("empty")
    df["BalanceValue"] = df["BalanceValue"].fillna("empty")
    
    


def lengthCorrectSturtzFormatConverter(length):
    
    length_string = str("%.2f" % round(length, 2)).replace(".", "")  # turn to string & remove decimal
    length_string = padFrontWithZeros(length_string, 5) # make it 5 characters
    
    return length_string


def padAssWithSpaces(inString, n):  # n is how many characters the field needs to be
    while len(inString) < n:
        inString = inString + " "

    return inString


def padFrontWithZeros(inString, n):  # n is how many characters the field needs to be
    while len(inString) < n:
        inString = "0" + inString

    return inString


def determineColor(inColor):
    if inColor == "White/White":
        return "WHWH"

    elif inColor == "Clay/Clay":
        return "CLCL"

    elif inColor == "Almond/Almond":
        return "ALAL"

    elif inColor == "Bronze/White":
        return "BZWH"

    elif inColor == "Bronze/Bronze":
        return "BZBZ"

    elif inColor == "Black/White":
        return "BKWH"

    elif inColor == "Black/Black":
        return "BKBK"

def detCutLength_Jambs(frameHeight):
    return round((frameHeight + 0.25), 2)


def detCutLength_HeadSill(frameWidth):
    return round((frameWidth + 0.25), 2)


def detCutLength_SashHorizontal(frameWidth):
    sashLength = frameWidth - 2.563 + 0.25
    return round(sashLength, 2)

def detCutLength_SashStileSlider(frameWidth):
    sashLength = (frameWidth/2.0) -1.375 + 0.25
    return round(sashLength, 2)

def detCutLength_SashRailSlider(frameHeight):
    sashLength = frameHeight - 2.875 + 0.25
    return round(sashLength, 2)


def detCutLength_SashVertical(frameHeight):
    sashLength = (frameHeight / 2.0) - 1.0 + 0.25
    return round(sashLength, 2)

def detProfileID_Jamb(wType, profile, nailfin, jLeg, flangeAddOn, channelFiller):
    profileID = ""

    if profile == "Bevel":
        profileID = profileID + "BV"
    elif profile == "Brickmould":
        profileID = profileID + "BM"

    if wType == "Slider":
        profileID = profileID + "S"
    elif wType == "Single Hung":
        profileID = profileID + "H"

    profileID = profileID + "L"  # L for jambs

    if nailfin == "Yes":
        profileID = profileID + "F"
    elif nailfin == "No":
        profileID = profileID + "N"

    if jLeg == "Yes":
        profileID = profileID + "J"
    elif jLeg == "No":
        profileID = profileID + "N"

    if flangeAddOn == "Yes":
        profileID = profileID + "S"
    elif flangeAddOn == "No":
        profileID = profileID + "N"

    if channelFiller == "Yes":
        profileID = profileID + "C"
    elif channelFiller == "No":
        profileID = profileID + "N"

    return profileID


def detProfileID_HeadSill(wType, profile, nailfin, jLeg, flangeAddOn, channelFiller):
    profileID = ""

    if profile == "Bevel":
        profileID = profileID + "BV"
    elif profile == "Brickmould":
        profileID = profileID + "BM"

    if wType == "Slider":
        profileID = profileID + "S"
    elif wType == "Single Hung":
        profileID = profileID + "H"

    profileID = profileID + "C"  # C for head/sill

    if nailfin == "Yes":
        profileID = profileID + "F"
    elif nailfin == "No":
        profileID = profileID + "N"

    if jLeg == "Yes":
        profileID = profileID + "J"
    elif jLeg == "No":
        profileID = profileID + "N"

    if flangeAddOn == "Yes":
        profileID = profileID + "S"
    elif flangeAddOn == "No":
        profileID = profileID + "N"

    if channelFiller == "Yes":
        profileID = profileID + "C"
    elif channelFiller == "No":
        profileID = profileID + "N"

    return profileID


def detProfileID_SashHorizontal(wType):
    profileID = ""

    if wType == "Slider":
        profileID = profileID + "SS"
        profileID = profileID + "H"
        profileID = profileID + "4"

    elif wType == "Single Hung":
        profileID = profileID + "SH"
        profileID = profileID + "R"
        profileID = profileID + "4"

    return profileID


def detProfileID_SashVertical(wType, config):
    profileID = ""

    if wType == "Slider":
        profileID = profileID + "SS"
        profileID = profileID + "V"

        if config == "RH":
            profileID = profileID + "R"
        else:  # making Left Hand the default
            profileID = profileID + "L"

        profileID = profileID + "4"

    elif wType == "Single Hung":
        profileID = profileID + "SH"
        profileID = profileID + "S"
        profileID = profileID + "4"

    return profileID


def generateWelderCode_Frame(profile, wType, nailfin, jLeg, flangeAddOn, channelFiller):
    welderCode = ""

    if profile == "Bevel":
        welderCode = welderCode + "V"
    elif profile == "Brickmould":
        welderCode = welderCode + "M"

    if wType == "Slider":
        welderCode = welderCode + "S"
    elif wType == "Single Hung":
        welderCode = welderCode + "H"

    if nailfin == "Yes":
        welderCode = welderCode + "F"
    elif nailfin == "No":
        welderCode = welderCode + "N"

    if jLeg == "Yes":
        welderCode = welderCode + "J"
    elif jLeg == "No":
        welderCode = welderCode + "N"

    if flangeAddOn == "Yes" and channelFiller == "No":
        welderCode = welderCode + "S"
    elif flangeAddOn == "No" and channelFiller == "Yes":
        welderCode = welderCode + "J"
    elif flangeAddOn == "Yes" and channelFiller == "Yes":
        welderCode = welderCode + "B"
    elif flangeAddOn == "No" and channelFiller == "No":
        welderCode = welderCode + "N"

    return welderCode


def generateWelderCode_Sash(wType, impact):
    welderCode = ""

    if wType == "Slider":
        welderCode = welderCode + "SSSTD"

    elif wType == "Single Hung":
        if impact == "Yes":
            welderCode = welderCode + "SHIMP"
        else:  # assuming default to be non-imact
            welderCode = welderCode + "SHSTD"

    return welderCode


def generateLine(orderNum, profileID, color, binNum, qty, comment, length):
    lineString = ""

    lineString = lineString + "K" + orderNum
    lineString = lineString + "P" + profileID
    lineString = lineString + "T" + color
    lineString = lineString + "N" + binNum
    lineString = lineString + "A" + qty
    lineString = lineString + "C" + comment
    lineString = lineString + "L" + length
    lineString = lineString + "\n"

    return lineString


def generateFrameLabelData(i,welderCode,frameHeight,frameWidth,orderNum,customer,
    destination,date,color,binNum,balanceType, CSL, sashHeight, scannableOrderNum):
    
    scannableOrderNum = str(scannableOrderNum)

    lineString = ""
    lineString = lineString + "LabId;" + str(i + 1) + ";\n"
    lineString = lineString + "FE=Sturtz1;\n"

    frameHeight_string = lengthCorrectSturtzFormatConverter(frameHeight)
    frameWidth_string = lengthCorrectSturtzFormatConverter(frameWidth)
    
    #lineString = lineString + "BA1=" + welderCode + ";\n"
    lineString = lineString + "BA1=" + welderCode + frameHeight_string + frameWidth_string + ";\n"

    #lineString = lineString + "BA2=" + orderNum + ";\n"
    lineString = lineString + "BA2=" + scannableOrderNum + ";\n"

    meetingRailLocation = frameHeight - (frameHeight / 2.0) - 0.231 - 0.125
    lineString = lineString + "BA3=" + str("%.3f" % round(meetingRailLocation, 3)) + ";\n"

    lineString = lineString + "S1=" + orderNum + ";\n"
    lineString = lineString + "S2=" + customer + ";\n"
    lineString = lineString + "S3=" + "###" + ";\n"
    
    if date != "###": #Got weird cause "date" is actually type "timestamp"
       lineString = lineString + "S4=" + str(date.date()) + ";\n" 
    else:
        lineString = lineString + "S4=" + date + ";\n"
        
    lineString = lineString + "S5=" + "###" + ";\n"
    lineString = lineString + "S6=" + destination + ";\n"
    lineString = lineString + "S7=" + "###" + ";\n"
    lineString = lineString + "S8=" + str("%.3f" % round(frameWidth, 3)) + "x" + str("%.3f" % round(frameHeight, 3)) + ";\n"
    lineString = lineString + "S9=" + color + ";\n"
    lineString = lineString + "S10=" + "###" + ";\n"
    lineString = lineString + "S11=" + "###" + ";\n"
    lineString = lineString + "S12=" + "###" + ";\n"
    lineString = lineString + "S13=" + "###" + ";\n"
    lineString = lineString + "S14=" + "###" + ";\n"
    lineString = lineString + "S15=" + "###" + ";\n"
    lineString = lineString + "S16=" + "###" + ";\n"
    lineString = lineString + "S17=" + str(detCutLength_Jambs(frameHeight)) + ";\n"
    lineString = lineString + "S18=" + str(detCutLength_HeadSill(frameWidth)) + ";\n"
    lineString = lineString + "S19=" + "###" + ";\n"
    lineString = lineString + "S20=" + "###" + ";\n"
    lineString = lineString + "S21=" + str("%.3f" % round((sashHeight), 3)) + ";\n"
    lineString = lineString + "S22=" + "###" + ";\n"
    lineString = lineString + "S23=" + "###" + ";\n"
    lineString = lineString + "S24=" + "###" + ";\n"
    lineString = lineString + "S25=" + "###" + ";\n"
    lineString = lineString + "S26=" + "###" + ";\n"
    lineString = lineString + "S27=" + "###" + ";\n"
    lineString = lineString + "S28=" + "###" + ";\n"
    lineString = lineString + "S29=" + "###" + ";\n"
    lineString = lineString + "S30=" + "###" + ";\n"
    lineString = lineString + "S31=" + "###" + ";\n"
    lineString = lineString + "S32=" + "###" + ";\n"
    lineString = lineString + "S33=" + "###" + ";\n"
    lineString = lineString + "S34=" + balanceType + ";\n"
    lineString = lineString + "S35=" + "###" + ";\n"
    lineString = lineString + "S36=" + CSL + ";\n"
    
    if CSL == "Yes":
        lineString = lineString + "S37=" + str("%.3f" % round((sashHeight+4.41), 3)) + ";\n"
        
    else:
        lineString = lineString + "S37=" + "###" + ";\n"
    
    lineString = lineString + "S38=" + "###" + ";\n"
    lineString = lineString + "S39=" + "###" + ";\n"
    lineString = lineString + "S40=" + "###" + ";\n"
    lineString = lineString + "S41=" + "###" + ";\n"
    lineString = lineString + "S42=" + "###" + ";\n"
    lineString = lineString + "S43=" + "###" + ";\n"
    lineString = lineString + "S44=" + "###" + ";\n"
    lineString = lineString + "S45=" + binNum + ";\n"
    lineString = lineString + "S46=" + "###" + ";\n"
    lineString = lineString + "\n\n"

    return lineString

def generateSashLabelData(i,welderCode,sashHeight,sashWidth,orderNum,customer,
    destination,date,color,binNum, frameWidth, frameHeight, scannableOrderNum, wasPair):

    scannableOrderNum = str(scannableOrderNum)
    lineString = ""
    if wasPair == False:  lineString = lineString + "LabId;" + str(i + 1) + ";\n"
   
    lineString = lineString + "FE=Sturtz1;\n"
    

    

    sashHeight_string = lengthCorrectSturtzFormatConverter(sashHeight)
    sashWidth_string = lengthCorrectSturtzFormatConverter(sashWidth)
    
    
    
   # lineString = lineString + "BA1=" + orderNum + ";\n"
    lineString = lineString + "BA1=" + scannableOrderNum + ";\n"
    #lineString = lineString + "testing normal: " + str(sashWidth) + ";\n"
   # lineString = lineString + "testing rounded: " + str("%.4f" % round(sashWidth, 4)) + ";\n"
   # lineString = lineString + "testing alreadyMadeSterintg: " + sashWidth_string + ";\n"

    lineString = lineString + "BA2=" + str("%.3f" % round(sashWidth, 3)) + ";" + str("%.3f" % round(sashHeight, 3)) + ";\n"
    lineString = lineString + "BA3=" + welderCode + sashHeight_string + sashWidth_string + ";\n"
    
    if date != "###": #Got weird cause "date" is actually type "timestamp"
       lineString = lineString + "S1=" + str(date.date()) + ";\n" 
    else:
        lineString = lineString + "S1=" + date + ";\n"
        
    #lineString = lineString + "S1=" + str(date) + ";\n"
    lineString = lineString + "S2=" + "###" + ";\n"
    lineString = lineString + "S3=" + destination + ";\n"
    lineString = lineString + "S4=" + "###" + ";\n"
    lineString = lineString + "S5=" + str("%.3f" % round(sashWidth, 3)) + "x" + str("%.3f" % round(sashHeight, 3)) + ";\n"
    lineString = lineString + "S6=" + "###" + ";\n"
    lineString = lineString + "S7=" + "###" + ";\n"
    lineString = lineString + "S8=" + "###" + ";\n"
    lineString = lineString + "S9=" + "###" + ";\n"
    lineString = lineString + "S10=" + "###" + ";\n"
    lineString = lineString + "S11=" + "###" + ";\n"
    lineString = lineString + "S12=" + "###" + ";\n"
    lineString = lineString + "S13=" + str("%.3f" % round(frameWidth, 3)) + "x" + str("%.3f" % round(frameHeight, 3)) + ";\n"
    lineString = lineString + "S14=" + color + ";\n"
    lineString = lineString + "S15=" + "###" + ";\n"
    lineString = lineString + "S16=" + "###" + ";\n"
    lineString = lineString + "S17=" + "###" + ";\n"
    lineString = lineString + "S18=" + "###" + ";\n"
    lineString = lineString + "S19=" + "###" + ";\n"
    lineString = lineString + "S20=" + binNum + ";\n"
    lineString = lineString + "S21=" + customer + ";\n"
    lineString = lineString + "\n\n"

    return lineString

def addMacro_BalanceHole(cutLength, numOfCoils):
    macroString = ""
    macroString = macroString + "Fab;BALANCEHOLE;"
    
    dimC = ( ((cutLength-.25)/2) +.231 +.125  )
    location = cutLength - dimC + 9.375 - (1.3*(numOfCoils - 1.0))
    
    macroString = macroString + lengthCorrectSturtzFormatConverter(location)
    ###########################################macroString = macroString + lengthCorrectSturtzFormatConverter(39.019)
    
    return macroString

def addMacro_CoilTakeout(cutLength):
    macroString = ""
    macroString = macroString + "Fab;COIL_TAKEOUT;" + lengthCorrectSturtzFormatConverter(cutLength)

    return macroString

def addMacro_BlockTackleTakeout(cutLength):
    macroString = ""
    macroString = macroString + "Fab;BT_TAKEOUT;" + lengthCorrectSturtzFormatConverter(cutLength)

    return macroString

def addMacro_InHoles(cutLength, sashHeight):#ASK OPIE HOW 
    macroString = ""
    macroString = macroString + "Fab;IN_HOLES;"
    macroString = macroString + lengthCorrectSturtzFormatConverter(5.88) + ";\n"
    
    if (cutLength - .25) > 41.75:
        macroString = macroString + "Fab;IN_HOLES;"
        #macroString = macroString + lengthCorrectSturtzFormatConverter(sashHeight - 1.25) + ";\n" //I think I originally got this from Opie's spreadsheet
        position = (cutLength / 2.0) - .231 - 2.144
        macroString = macroString + lengthCorrectSturtzFormatConverter(position) + ";\n"
    
    macroString = macroString + "Fab;IN_HOLES;"
    macroString = macroString + lengthCorrectSturtzFormatConverter(cutLength - 5.88) 
    
    return macroString

def addMacro_BTHole(cutLength):
    macroString = ""
    
    macroString = macroString + "Fab;BT_HOLES;"
    macroString = macroString + lengthCorrectSturtzFormatConverter(cutLength - 1.125 - .5) 

    return macroString

def addMacro_NightLatch(sashHeight): #Should be same as child safety latch
    macroString = ""
    
    macroString = macroString + "Fab;NIGHTLATCHLEFT;"
    #macroString = macroString + lengthCorrectSturtzFormatConverter(sashHeight + 4.41) //I think I had originally gotten this from Opie's spreadsheet
    position = .7190 + sashHeight + 4.375 + (1.875/2)

    macroString = macroString + lengthCorrectSturtzFormatConverter(position) 

    return macroString

def addMacro_Mullion(location): 
    macroString = ""
    
    macroString = macroString + "Fab;MULLION;"
    macroString = macroString + lengthCorrectSturtzFormatConverter(location) 

    return macroString


def detNumOfCoils(balance):
    
    numOfCoils = 0
    
    if balance == 35:
        numOfCoils = 1
        
    elif balance == 45:
        numOfCoils = 1
        
    elif balance == 55:
        numOfCoils = 1
        
    elif balance == 65:
        numOfCoils = 1
        
    elif balance == 75:
        numOfCoils = 1

    elif balance == 90:
        numOfCoils = 1

    elif balance == 100:
        numOfCoils = 1
        
    elif balance == 110:
        numOfCoils = 2

    elif balance == 120:
        numOfCoils = 2
        
    elif balance == 130:
        numOfCoils = 2
        
    elif balance == 140:
        numOfCoils = 2
        
    elif balance == 150:
        numOfCoils = 2
        
    elif balance == 165:
        numOfCoils = 2
        
    elif balance == 180:
        numOfCoils = 2
        
    elif balance == 190:
        numOfCoils = 2
        
    elif balance == 200:
        numOfCoils = 2
        
    elif balance == 215:
        numOfCoils = 3
        
    elif balance == 225:
        numOfCoils = 3
        
    elif balance == 245:
        numOfCoils = 3
        
    elif balance == 255:
        numOfCoils = 3
        
    elif balance == 270:
        numOfCoils = 3
        
    elif balance == 280:
        numOfCoils = 3
        
    elif balance == 290:
        numOfCoils = 3
        
    elif balance == 300:
        numOfCoils = 3
        
    elif balance == 315:
        numOfCoils = 4
        
    elif balance == 330:
        numOfCoils = 4
        
    elif balance == 345:
        numOfCoils = 4
        
    elif balance == 360:
        numOfCoils = 4
        
    elif balance == 370:
        numOfCoils = 4
        
    elif balance == 380:
        numOfCoils = 4
        
    elif balance == 390:
        numOfCoils = 4
        
    elif balance == 400:
        numOfCoils = 4
        
    else: #defaulting to 1
        numOfCoils = 1    

    return numOfCoils

def addMacro_tapconJamb():
    hi = "Hi"
    ## too lazy to reverse engineer the tapcon logic lol
    return hi

def addFabMacros_Jamb(balanceType, childSafetyLatch, cutLength, sashHeight, nailfin, flangeAddOn, balance):
    
    macroString = ""
    
    if balanceType == "Coil":
        
        numOfCoils = detNumOfCoils(balance)
        
        macroString = macroString + addMacro_CoilTakeout(cutLength) + ";\n"
        macroString = macroString + addMacro_BalanceHole(cutLength, numOfCoils) + ";\n"
       
        
    elif balanceType == "Block&Tackle":
        macroString = macroString + addMacro_BlockTackleTakeout(cutLength) + ";\n"
        macroString = macroString + addMacro_BTHole(cutLength) + ";\n"
        
    if childSafetyLatch == "Yes":
        macroString = macroString + addMacro_NightLatch(sashHeight) + ";\n"
        
    if nailfin == "No" and flangeAddOn == "No":
        macroString = macroString + addMacro_InHoles(cutLength, sashHeight) + ";\n"
        
    if flangeAddOn == "Yes":
        macroString = macroString + addMacro_tapconJamb() + ";\n"

    file_Jamb.write(macroString)
    

    
def addFabMacros_HeadSill(mullion, cutLength):
    
    if mullion != "Single":
        macroString = ""

        if mullion == "Twin":
            macroString = macroString + addMacro_Mullion(cutLength / 2.0) + ";\n"
        
        elif mullion == "Triple":
            sidelite = ((cutLength - .25) - 1) / 3
            dimC = sidelite - .698 + .125
            macroString = macroString + addMacro_Mullion(.434 + dimC + (.172/2) + (1.290/2)) + ";\n"
            macroString = macroString + addMacro_Mullion( cutLength - .434 - dimC - (.172/2) - (1.290/2)  ) + ";\n"

        file_HeadSill.write(macroString)

def detFileNameParameters():
    fileNameDate = ""
    counter = "1"
    
#####################stuff for filename
    if df.iloc[0]["Schedule Date"] == "###":
        fileNameDate = today
    
    else:
        fileNameDate = df.iloc[0]["Schedule Date"].strftime("%Y-%m-%d")
        
#######################stuff for counter
    countConfig = "countConf.jo"

    if not os.path.exists(countConfig): #make config file if not there
        countConfigFile = open("countConf.jo", "w")
        countConfigFile.write("Date: " + today + "\n")
        countConfigFile.write("Count: " + "1" + "\n")
        countConfigFile.close()
        

    #get current values from config file
    countConfigFile = open("countConf.jo", "r")
    currentDateInConfig = countConfigFile.readline()
    currentCountInConfig = countConfigFile.readline()

    currentDateInConfig = currentDateInConfig.replace("Date: ", "")
    currentCountInConfig = currentCountInConfig.replace("Count: ", "")
    currentDateInConfig = currentDateInConfig.replace("\n", "")
    currentCountInConfig = currentCountInConfig.replace("\n", "")

    countConfigFile.close()
    

    #Update parameters in config file
    if currentDateInConfig != fileNameDate:
        currentDateInConfig = fileNameDate
        currentCountInConfig = "1"
        
    counter = currentCountInConfig
    currentCountInConfig = str(int(currentCountInConfig)+1)

    countConfigFile = open("countConf.jo", "w")
    countConfigFile.write("Date: " + currentDateInConfig + "\n")
    countConfigFile.write("Count: " + currentCountInConfig + "\n")
    countConfigFile.close()
     

    return fileNameDate, counter
    

cleanDataframe()



fileNameDate, fileCounter = detFileNameParameters()

file_Jamb = open(fileCounter + " JambFrameSaw-" + fileNameDate + ".SAW", "w")
file_HeadSill = open(fileCounter + " HeadSillFrameSaw-" + fileNameDate + ".SAW", "w")
file_HorizontalSash = open(fileCounter + " LockLiftRailSashSaw-" + fileNameDate + ".SAW", "w")
file_VerticalSash = open(fileCounter + " StileSashSaw-" + fileNameDate + ".SAW", "w")

file_JambLabel = open(fileCounter + " JambFrameSaw-" + fileNameDate + ".la1", "w")
file_HeadSillLabel = open(fileCounter + " HeadSillFrameSaw-" + fileNameDate + ".la1", "w")
file_HorizontalSashLabel = open(fileCounter + " LockLiftRailSashSaw-" + fileNameDate + ".la1", "w")
file_VerticalSashLabel = open(fileCounter + " StileSashSaw-" + fileNameDate + ".la1", "w")

prevUsed = False #used for writing info as pair for sashes
sashCounter = 0 #used for the sash .la1 files 
wasSashPair = False #you guessed it, used for sash stuff.  Didn't need this, turned out to be the same as prevUsed

for i, rows in df.iterrows():

    orderNum = str(df.iloc[i]["Order Number"])
    #orderNum = df.iloc[i]["Order Number"]
    scannableOrderNum = df.iloc[i]["Full Scannable Order Number"]
    binNum = "Z00Z"  # Using a nonsense value for now
    qty = "001"
    comment = orderNum
    balance = df.iloc[i]["BalanceValue"]
    

    if scannableOrderNum == "empty":
        scannableOrderNum = orderNum

    color = determineColor(df.iloc[i]["Color"])
    cutLength_jamb = detCutLength_Jambs(df.iloc[i]["Frame Height"])

    cutLength_headSill = detCutLength_HeadSill(df.iloc[i]["Frame Width"])
    
    if df.iloc[i]["Window Type"] == "Slider":
        cutLength_sashHorizontal = detCutLength_SashRailSlider(df.iloc[i]["Frame Height"])
        cutLength_sashVertical = detCutLength_SashStileSlider(df.iloc[i]["Frame Width"])  
        
        profileID_sashVertical = detProfileID_SashHorizontal(df.iloc[i]["Window Type"])
        profileID_sashHorizontal = detProfileID_SashVertical(
            df.iloc[i]["Window Type"], df.iloc[i]["Configuration (for sliders)"]
    )

    else:
        cutLength_sashHorizontal = detCutLength_SashHorizontal(df.iloc[i]["Frame Width"])
        cutLength_sashVertical = detCutLength_SashVertical(df.iloc[i]["Frame Height"])
        profileID_sashHorizontal = detProfileID_SashHorizontal(df.iloc[i]["Window Type"])
        profileID_sashVertical = detProfileID_SashVertical(
             df.iloc[i]["Window Type"], df.iloc[i]["Configuration (for sliders)"]
        )
    
    
    profileID_Jamb = detProfileID_Jamb(
        df.iloc[i]["Window Type"],
        df.iloc[i]["Frame Profile"],
        df.iloc[i]["Nailfin?"],
        df.iloc[i]["J-leg?"],
        df.iloc[i]["Flange add-on?"],
        df.iloc[i]["Channel Filler?"]
    )
    profileID_HeadSill = detProfileID_HeadSill(
        df.iloc[i]["Window Type"],
        df.iloc[i]["Frame Profile"],
        df.iloc[i]["Nailfin?"],
        df.iloc[i]["J-leg?"],
        df.iloc[i]["Flange add-on?"],
        df.iloc[i]["Channel Filler?"]
    )
    welderCode_frame = generateWelderCode_Frame(
        df.iloc[i]["Frame Profile"],
        df.iloc[i]["Window Type"],
        df.iloc[i]["Nailfin?"],
        df.iloc[i]["J-leg?"],
        df.iloc[i]["Flange add-on?"],
        df.iloc[i]["Channel Filler?"]
    )
    
    
    
    welderCode_sash = generateWelderCode_Sash(
        df.iloc[i]["Window Type"], df.iloc[i]["Impact?"]
    )

    # pad values to correct character count for the field
    orderNum = padAssWithSpaces(orderNum, 10)
    #comment = padAssWithSpaces(comment, 10)
    comment = padAssWithSpaces(comment, 60)
    profileID_sashHorizontal = padAssWithSpaces(profileID_sashHorizontal, 10)
    profileID_sashVertical = padAssWithSpaces(profileID_sashVertical, 10)
    profileID_Jamb = padAssWithSpaces(profileID_Jamb, 10)
    profileID_HeadSill = padAssWithSpaces(profileID_HeadSill, 10)
    
    # convert lengths to strings
    cutLength_jamb_string = lengthCorrectSturtzFormatConverter(cutLength_jamb)
    cutLength_headSill_string = lengthCorrectSturtzFormatConverter(cutLength_headSill)
    cutLength_sashHorizontal_string = lengthCorrectSturtzFormatConverter(cutLength_sashHorizontal)
    cutLength_sashVertical_string = lengthCorrectSturtzFormatConverter(cutLength_sashVertical)

    # write the lines to the .SAW files for frame
    file_Jamb.write( generateLine(orderNum, profileID_Jamb, color, binNum, qty, comment, cutLength_jamb_string) )
    file_HeadSill.write( generateLine(orderNum, profileID_HeadSill, color, binNum, qty, comment, cutLength_headSill_string) )
    
    # write the lines to the .SAW files for sash

    if i != 0:
        
        if prevUsed == False:
            if (df.iloc[i]["Frame Height"] == df.iloc[i-1]["Frame Height"]) and (df.iloc[i]["Frame Width"] == df.iloc[i-1]["Frame Width"]): ##send as pair
                pairComment = str(df.iloc[i-1]["Order Number"]) + ";" + str(df.iloc[i]["Order Number"])
                pairComment = padAssWithSpaces(pairComment, 60)
                file_HorizontalSash.write( generateLine(orderNum, profileID_sashHorizontal, color, binNum, qty, pairComment , cutLength_sashHorizontal_string) )
                file_VerticalSash.write( generateLine(orderNum, profileID_sashVertical, color, binNum, qty, pairComment, cutLength_sashVertical_string) )
               # wasSashPair = True
                prevUsed = True

                #This is gunna be nasty regenerating previous info cause this addition was an afterthought and I'm being lazy and don't wanna
#rewrite code to make this prettier
            else:
                prevOrderNum = str(df.iloc[i-1]["Order Number"])
                prevOrderNum = padAssWithSpaces(prevOrderNum, 10)
                prevProfileID_sashHorizontal = detProfileID_SashHorizontal(df.iloc[i-1]["Window Type"])
                prevProfileID_sashHorizontal = padAssWithSpaces(prevProfileID_sashHorizontal, 10)
                prevColor = determineColor(df.iloc[i-1]["Color"])
                prevComment = padAssWithSpaces(prevOrderNum, 60)
                prevCutLength_sashHorizontal = detCutLength_SashHorizontal(df.iloc[i-1]["Frame Width"])
                prevcutLength_sashHorizontal_string = lengthCorrectSturtzFormatConverter(prevCutLength_sashHorizontal)
                prevprofileID_sashVertical = detProfileID_SashVertical(
             df.iloc[i-1]["Window Type"], df.iloc[i-1]["Configuration (for sliders)"])
        
                prevprofileID_sashVertical = padAssWithSpaces(prevprofileID_sashVertical, 10)
                prevcutLength_sashVertical = detCutLength_SashVertical(df.iloc[i-1]["Frame Height"])
                prevcutLength_sashVertical_string = lengthCorrectSturtzFormatConverter(prevcutLength_sashVertical)
                
                file_HorizontalSash.write( generateLine(prevOrderNum, prevProfileID_sashHorizontal, prevColor, binNum, qty, prevComment, prevcutLength_sashHorizontal_string) )
                file_VerticalSash.write( generateLine(prevOrderNum, prevprofileID_sashVertical, prevColor, binNum, qty, prevComment, prevcutLength_sashVertical_string) )
            
        else: prevUsed = False
            


    # write macro stuff into the .SAW files
    addFabMacros_Jamb(df.iloc[i]["Balance Type"], df.iloc[i]["Child Safety Latch"], cutLength_jamb, (cutLength_sashVertical - .25), df.iloc[i]["Nailfin?"], df.iloc[i]["Flange add-on?"], balance )
    addFabMacros_HeadSill(df.iloc[i]["Mullion"], cutLength_headSill)



    # write data into the .la1 files

    file_JambLabel.write( generateFrameLabelData(i, welderCode_frame, df.iloc[i]["Frame Height"], df.iloc[i]["Frame Width"],
            orderNum, df.iloc[i]["Customer"], df.iloc[i]["Destination"], df.iloc[i]["Schedule Date"], color, binNum, df.iloc[i]["Balance Type"],
            df.iloc[i]["Child Safety Latch"], (cutLength_sashVertical - .25), scannableOrderNum) )
    file_HeadSillLabel.write( generateFrameLabelData(i, welderCode_frame, df.iloc[i]["Frame Height"], df.iloc[i]["Frame Width"],
            orderNum, df.iloc[i]["Customer"], df.iloc[i]["Destination"], df.iloc[i]["Schedule Date"], color, binNum, df.iloc[i]["Balance Type"],
            df.iloc[i]["Child Safety Latch"], (cutLength_sashVertical - .25), scannableOrderNum) )
    


    if df.iloc[i]["Window Type"] == "Slider":
        file_HorizontalSashLabel.write( generateSashLabelData(sashCounter, welderCode_sash, (cutLength_sashHorizontal - .25), 
                (cutLength_sashVertical - .25), orderNum, df.iloc[i]["Customer"], df.iloc[i]["Destination"], 
                df.iloc[i]["Schedule Date"], color, binNum, df.iloc[i]["Frame Width"], df.iloc[i]["Frame Height"], scannableOrderNum, prevUsed) )
        file_VerticalSashLabel.write( generateSashLabelData(sashCounter, welderCode_sash, (cutLength_sashHorizontal - .25), 
                (cutLength_sashVertical - .25), orderNum, df.iloc[i]["Customer"], df.iloc[i]["Destination"], 
                df.iloc[i]["Schedule Date"], color, binNum, df.iloc[i]["Frame Width"], df.iloc[i]["Frame Height"], scannableOrderNum, prevUsed)  )

    else:
        file_HorizontalSashLabel.write( generateSashLabelData(sashCounter, welderCode_sash, (cutLength_sashVertical - .25), 
                (cutLength_sashHorizontal - .25), orderNum, df.iloc[i]["Customer"], df.iloc[i]["Destination"], 
                df.iloc[i]["Schedule Date"], color, binNum, df.iloc[i]["Frame Width"], df.iloc[i]["Frame Height"], scannableOrderNum, prevUsed) )
        file_VerticalSashLabel.write( generateSashLabelData(sashCounter, welderCode_sash, (cutLength_sashVertical - .25), 
                (cutLength_sashHorizontal - .25), orderNum, df.iloc[i]["Customer"], df.iloc[i]["Destination"], 
                df.iloc[i]["Schedule Date"], color, binNum, df.iloc[i]["Frame Width"], df.iloc[i]["Frame Height"], scannableOrderNum, prevUsed) )
        
    if prevUsed == False: sashCounter = sashCounter + 1

#this gets the last value, needed to be outside of the for loop
if prevUsed == False:
    prevOrderNum = str(df.iloc[-1]["Order Number"])
    prevOrderNum = padAssWithSpaces(prevOrderNum, 10)
    prevProfileID_sashHorizontal = detProfileID_SashHorizontal(df.iloc[i-1]["Window Type"])
    prevProfileID_sashHorizontal = padAssWithSpaces(prevProfileID_sashHorizontal, 10)
    prevColor = determineColor(df.iloc[-1]["Color"])
    prevComment = padAssWithSpaces(prevOrderNum, 60)
    prevCutLength_sashHorizontal = detCutLength_SashHorizontal(df.iloc[-1]["Frame Width"])
    prevcutLength_sashHorizontal_string = lengthCorrectSturtzFormatConverter(prevCutLength_sashHorizontal)
    prevprofileID_sashVertical = detProfileID_SashVertical(
    df.iloc[-1]["Window Type"], df.iloc[-1]["Configuration (for sliders)"])
        
    prevprofileID_sashVertical = padAssWithSpaces(prevprofileID_sashVertical, 10)
    prevcutLength_sashVertical = detCutLength_SashVertical(df.iloc[i-1]["Frame Height"])
    prevcutLength_sashVertical_string = lengthCorrectSturtzFormatConverter(prevcutLength_sashVertical)
                
    file_HorizontalSash.write( generateLine(prevOrderNum, prevProfileID_sashHorizontal, prevColor, binNum, qty, prevComment, prevcutLength_sashHorizontal_string) )
    file_VerticalSash.write( generateLine(prevOrderNum, prevprofileID_sashVertical, prevColor, binNum, qty, prevComment, prevcutLength_sashVertical_string) )
    

file_Jamb.close()
file_HeadSill.close()
file_HorizontalSash.close()
file_VerticalSash.close()

file_JambLabel.close()
file_HeadSillLabel.close()
file_HorizontalSashLabel.close()
file_VerticalSashLabel.close()

print("PO download files generated! You're welcome")

time.sleep(3)

print("Byeee")

time.sleep(1)