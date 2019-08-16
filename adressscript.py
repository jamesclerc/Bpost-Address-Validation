import requests
import pandas as pd
import json

print("script starting")

print("loading Excels")

dfs = pd.read_excel("./AddressData.xlsx")
dfs["Error"] = "false"
dfs["Warning"] = "false"

headers = {
    "Content-Type" : "application/json"
}

x = 0

error = []
warning = []
warning_componentRef = []
error_componentRef = []
global_result = []

fil3 = open("result.json", "w")

while x < len(dfs["Address"]):

    body = {
        "ValidateAddressesRequest":{
            "AddressToValidateList":{
                "AddressToValidate":[{
                    "@id":str(x),
                    "AddressBlockLines":{
                        "UnstructuredAddressLine": [{
                                "@locale" : "fr",
                                "*body" :  str(dfs["Address"][x]).replace(",", "") + " " + str(dfs["Zip Code"][x]) + " " + str(dfs["City"][x]) + " "  + str(dfs["Country"][x].upper())
                        }]
                    }
                }]
            }
        }
    }
    req = requests.post("https://webservices-pub.bpost.be/ws/ExternalMailingAddressProofingCSREST_v1/address/validateAddresses", headers=headers, json=body)
    print(req.status_code)
    while (req.status_code != 200):
        req = requests.post("https://webservices-pub.bpost.be/ws/ExternalMailingAddressProofingCSREST_v1/address/validateAddresses", headers=headers, json=body)
    #print (req.text)
    result = json.loads(req.text)
    global_result.append(result)
    #print(result)
    # for x in len(result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"]):
    #     print(x)
    print(result)
    print("\n\n")
    fil3.write(json.dumps(result))
    fil3.write("\n")
    if ("Error" in result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]):
        for y in range(len(result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"])):
            if ("error" in result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][y]["ErrorSeverity"]):
                # if (x == 0):
                #     print("test")
                #     print(dfs["Error"][x])
                dfs["Error"][x] = "true"
                #print(dfs["Error"][x])
                if (result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][y]["ErrorCode"] not in error):
                    error.append(str(x) + "-->" + result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][y]["ErrorCode"])
                if (result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][y]["ComponentRef"] not in error_componentRef):
                    error.append(str(x) + "-->" + result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][y]["ComponentRef"] + " | " + result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][y]["ErrorCode"])
                #print("ERREUR")
            if ("warning" in result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][y]["ErrorSeverity"]):
                dfs["Warning"][x] = "true"
                if (result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][y]["ErrorCode"] not in warning):
                    warning.append(str(x) + "-->" + result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][y]["ErrorCode"])
                if (result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][y]["ComponentRef"] not in warning_componentRef):
                    warning.append(str(x) + "-->" +result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][y]["ComponentRef"] + " | " + result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][y]["ErrorCode"])
                #print("WARNING")
    

            #print(result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][x]["ErrorSeverity"])
    # if (result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"][0]["ErrorSeverity"] == "error"):
    #     print("Address is " + adress + json.dumps(result["ValidateAddressesResponse"]["ValidatedAddressResultList"]["ValidatedAddressResult"][0]["Error"]))
    # else:
    #     print("Address is " + adress)
    # if (x % 200 == 0):
#    if(x == 100):
#        break
        #export_excel = dfs.to_excel("./Return2.xlsx", index = None, header=True)
    x = x + 1

#export_excel = dfs.to_excel ("./Return2.xlsx", index = None, header=True)
fil = open("warning", "w")
for item in warning:
    fil.write(json.dumps(item))
    fil.write("\n")

fil2 = open("error", "w")
for item in error:
    fil2.write(json.dumps(item))
    fil2.write("\n")