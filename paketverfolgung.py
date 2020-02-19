from zeep import Client
import csv
import time
import glob
import ctypes

csv_files = glob.glob('*.csv')
dpd_csv = []
with open(csv_files[0]) as csvDataFile:
    csvReader = csv.reader(csvDataFile, delimiter=';')
    for row in csvReader:
        dpd_csv.append(row)

#print(str(dpd_csv))

# Sandbox daten DPD
delisId = "sandboxdpd" # "sandboxdpd"
password = "xMmshh1" # "xMmshh1"
lang = "de_DE"
paketId = "01405400945058"

client = Client(wsdl="https://public-ws-stage.dpd.com/services/LoginService/V2_0/?wsdl")
auth = client.service.getAuth(delisId, password, lang)

header_value = {
    'delisId': delisId,
    'authToken': auth['authToken'],
    'messageLanguage': lang
}

for line in dpd_csv:
    # paketId einspeichern
    paketId = line[0]

    time.sleep(1)

    if paketId.startswith("0134"):

        parcel = Client(wsdl="https://public-ws-stage.dpd.com/services/ParcelLifeCycleService/V2_0/?wsdl")

        parcel_info = parcel.service.getTrackingData(paketId, _soapheaders={'authentication': header_value})

        status_info = parcel_info['statusInfo']
        verlauf = []
        for status in status_info:
            verlauf.append(status['status'])

        last_status = verlauf[len(verlauf)-1]
        line.append(last_status)
        print(str(line))


with open('04.02.2020 dpd.csv', 'w', newline='') as f:
    writer = csv.writer(f, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    for row in dpd_csv:
        writer.writerow(row)

print(' ')
print(' ')
print('##################################')
print('#            FERTIG              #')
print('#            ~*~*~*~             #')
print('##################################')
ctypes.windll.user32.MessageBoxW(0, "Die Datei wurde mit allen Statusmeldungen versehen", "Status update", 1)
