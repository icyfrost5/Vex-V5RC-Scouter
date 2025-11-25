# Vex-scouter
Grabs Vex (any program in robotevents.com so V5RC, VIQ, Etc.) Teams id's, Team numbers, Team names, Organization, Grade level, Location, Highest skills score, Higest Driver score, Highest Programming score, Awards, and Best rank in a tournement all with just a robotevents.com link. 

all I kindly ask is for credit if you plan to copy this


# In Main.py edit the {INSERT YOUR API KEY HERE} with your robotevents.com api key open the requirements downloader and then open start.bat
- ps. don't do a lot in quick succession or you're going to get rate limited and I didn't add handling for it. (error at edata/tdata where it fails to get the json)


# How to use (assuming you have set it up correctly):

- enter the event link from robotevents.com (can be any subpage of the event)
- choose the name you want the Excel file to be named
- wait
- It gathers all the data for you and organizes it into an excel file in the same directory as main.py
