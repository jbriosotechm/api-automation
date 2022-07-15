import requests,os,sys
import dynamicConfig
import traceback
import json
import userConfig, SystemConfig
import time
sys.path.append("customLib")
import customLib.customLogging as customLogging

def formatRequestBody():
    if "application/x-www-form-urlencoded" == dynamicConfig.currentContentType:
        requestString = dynamicConfig.currentRequest.encode('ascii', 'ignore')
        if not requestString.startswith("{"):
            requestString = "{" + requestString
        if not requestString.endswith("}"):
            requestString = requestString + "}"
        requestString = requestString.replace("\n", ",")
        body = json.loads(requestString)
    else:
        body = dynamicConfig.currentRequest
    return body

def triggerSoapRequest():
    headers = dynamicConfig.currentHeader
    url     = dynamicConfig.currentUrl
    body    = dynamicConfig.currentRequest

    dynamicConfig.responseStatusCode = None
    dynamicConfig.responseHeaders    = None
    dynamicConfig.responseText       = None

    response = None

    try:
        response = requests.post(url,data=body,headers=headers,timeout=userConfig.timeoutInSeconds,verify=False)

        requestContent="URL : {0}\nHeaders : {1}\nBody: {2}".format(url,headers,body)
        customLogging.writeToLog("Req_SOAP_"+str(time.time()),requestContent)

    except Exception,e:
        traceback.print_exc()
        dynamicConfig.currentException=traceback.format_exc()

    dynamicConfig.currentResponse=response

    if response is not None:
        dynamicConfig.responseHeaders=response.headers
        dynamicConfig.responseStatusCode=response.status_code
        dynamicConfig.responseText=response.text

    print "\n*************** [ Response ] ***************"
    print "\n\n Headers : {0}".format(dynamicConfig.responseHeaders)
    print "\nStatus Code : {0}".format(dynamicConfig.responseStatusCode)
    print "\nBody : {0}".format(dynamicConfig.responseText)


    responseContent="Status Code : {0}\n\nHeaders : {1}\n\nBody : {2}".format(dynamicConfig.responseStatusCode,dynamicConfig.responseHeaders,dynamicConfig.responseText)
    customLogging.writeToLog("Res_SOAP_"+str(time.time()),responseContent)

def triggerRestRequest():
    headers        = dynamicConfig.currentHeader
    url            = dynamicConfig.currentUrl
    body           = formatRequestBody()
    requestType    = dynamicConfig.restRequestType
    authentication = dynamicConfig.currentAuthentication
    cookies        = dynamicConfig.currentCookie
    files          = dynamicConfig.currentfile
    timeout        = userConfig.timeoutInSeconds
    runTimes       = 1

    print "Request type is : ",requestType
    print "Request headers is : ",headers

    response=None
    if dynamicConfig.currentRequest is None:
        dynamicConfig.currentRequest = ""

    requestContent = "\nRequest type : {3}\n\nURL : {0}\n\nHeaders : {1}\n\nBody: {2}".format(url,headers,body,requestType)
    customLogging.writeToLog("Req_Rest_"+str(time.time()),requestContent)

    if "RERUN_TIMES" in SystemConfig.localRequestDict.keys():
        runTimes = int(SystemConfig.localRequestDict["RERUN_TIMES"])

    for i in range(0, runTimes):
        try:
            if str(requestType).startswith("post") and files is None:
                response = requests.post(url,data=body,headers=headers,timeout=timeout,verify=False, auth=authentication, cookies=cookies)
            elif str(requestType).startswith("post") and files is not None:
                response = requests.post(url,files=files,headers=headers,timeout=timeout,verify=False, auth=authentication, cookies=cookies)
            elif str(requestType).startswith("put"):
                response = requests.put(url,data=body,headers=headers,timeout=timeout,verify=False, auth=authentication, cookies=cookies)
            elif str(requestType).startswith("get"):
                response = requests.get(url,data=body,headers=headers,timeout=timeout,verify=False, auth=authentication, cookies=cookies)
            elif str(requestType).startswith("patch"):
                response = requests.patch(url,data=body,headers=headers,timeout=timeout,verify=False, auth=authentication, cookies=cookies)
            elif str(requestType).startswith("delete"):
                response = requests.delete(url,data=body,headers=headers,timeout=timeout,verify=False, auth=authentication, cookies=cookies)
            else:
                response = requests.post(url,data=body,headers=headers,timeout=timeout,verify=False, auth=authentication, cookies=cookies)

        except Exception,e:
            traceback.print_exc()
            dynamicConfig.currentException = traceback.format_exc()

    #response=response.decode("utf-8")
    dynamicConfig.currentResponse = response

    if response is not None:
        dynamicConfig.responseHeaders    = response.headers
        dynamicConfig.responseStatusCode = response.status_code
        dynamicConfig.responseText       = response.text.encode('ascii', 'ignore')

    print "\n*************** [ Response ] ***************"
    print "\n\n Headers : {0}".format(dynamicConfig.responseHeaders)
    print "\nStatus Code : {0}".format(dynamicConfig.responseStatusCode)

    if dynamicConfig.responseHeaders is not None:
        if "application/pdf" not in str(dynamicConfig.responseHeaders):
            print "\nBody : {0}".format(dynamicConfig.responseText)
            responseContent="Status Code : {0}\n\nHeaders : {1}\n\nBody : {2}".format(dynamicConfig.responseStatusCode,dynamicConfig.responseHeaders,dynamicConfig.responseText)
        else:
            print "\nBody : {0}".format(dynamicConfig.responseText)
            responseContent="Status Code : {0}\n\nHeaders : {1}".format(dynamicConfig.responseStatusCode,dynamicConfig.responseHeaders)

            pdfLocation = "response.pdf"
            if "PDF_LOCATION" in SystemConfig.localRequestDict.keys():
                pdfLocation = SystemConfig.localRequestDict["PDF_LOCATION"]
            with open(pdfLocation, 'wb') as f:
                f.write(dynamicConfig.currentResponse.content)

    customLogging.writeToLog("Res_Rest" + str(time.time()),responseContent)
