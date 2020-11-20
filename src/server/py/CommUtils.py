import json
from PythonAPI import sendSocketMessage

# def sendMessage(message, messageType="log"):
#     messageDict = {}
#     if messageType == "log":
#         messageDict["type"] = "log"
#         messageDict["message"] = message
#         print(json.dumps(messageDict), flush=True)
#     elif messageType == "prog":
#         messageDict["type"] = "prog"
#         messageDict["message"] = message
#         print(json.dumps(messageDict), flush=True)
#     elif messageType == "progMsg":
#         messageDict["type"] = "progMsg"
#         messageDict["message"] = message
#         print(json.dumps(messageDict), flush=True)
#     elif messageType == "requireInput":
#         messageDict["type"] = "requireInput"
#         messageDict["message"] = message
#         print(json.dumps(messageDict), flush=True)
#     else:
#         print(json.dumps("Failed Python message"), flush=True)

def sendMessage(message, messageType="log"):
    messageDict = {}
    if messageType == "log":
        messageDict["type"] = "log"
        messageDict["message"] = message
        sendSocketMessage("log", messageDict)
    elif messageType == "prog":
        messageDict["type"] = "prog"
        messageDict["message"] = message
        sendSocketMessage("prog", messageDict)
    elif messageType == "progMsg":
        messageDict["type"] = "progMsg"
        messageDict["message"] = message
        sendSocketMessage("progMsg", messageDict)
    elif messageType == "requireInput":
        messageDict["type"] = "requireInput"
        messageDict["message"] = message
        sendSocketMessage("requireInput", messageDict)
    else:
        sendSocketMessage("log", "Failed Python message")
