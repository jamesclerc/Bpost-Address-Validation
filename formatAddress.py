import json
import pandas as pd

data = open("./result.json", "r")

dfs = pd.read_excel("./AddressData.xlsx")

dfs["ValidAddress"] = "true"

x = 0
for line in data:
    print(x)
    resultLine = json.loads(line)
    resultLine = resultLine["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"]
    if ("Error" in resultLine[0]):
        for error in range(len(resultLine[0]["Error"])):
            if ("error" in resultLine[0]["Error"][error]["ErrorSeverity"]):
                dfs["ValidAddress"][x] = "false"
            if ("warning" in resultLine[0]["Error"][error]["ErrorSeverity"]):
                if (resultLine[0]["Error"][error]["ComponentRef"] == "UnstructuredDeliveryPointLocation" or resultLine[0]["Error"][error]["ComponentRef"] == "BoxNumber" or resultLine[0]["Error"][error]["ComponentRef"] == "MunicipalityName" or resultLine[0]["Error"][error]["ComponentRef"] == "PostalCode"):
                    print("minor warning")
                else:
                    print("major warning")
                    dfs["ValidAddress"][x] = "false"
    x = x + 1

p =  dfs.to_excel("./final.xlsx", index=None, header=True)