from flask import Flask, request
from engineio.async_drivers import eventlet
from flask_socketio import SocketIO, send, emit
import sys, json
import ElentraReport as ER
import IntervalTimer as inter
from datetime import datetime

# The following pages were helpful in debugging the packaging
# of flask-socketio:
#   https://stackoverflow.com/questions/63254533/why-does-pyinstaller-fail-to-package-eventlet
#   https://github.com/pyinstaller/pyinstaller/issues/2572
#   https://pyinstaller.readthedocs.io/en/stable/hooks.html

app = Flask(__name__)
socketio = SocketIO(app, 
    async_mode='eventlet', 
    cors_allowed_origins="*", 
    engineio_logger=True,
    ping_timeout=300,
    ping_interval=30)


@socketio.on('connect')
def handleConnection():
    send("Electron has connected")

# Called by comms.sendMessage after it formats the message
def sendSocketMessage(channel, msg):
    send("sendSocketMessage called!")
    emit(channel, json.dumps(msg))

def socketSleep():
    socketio.sleep(0)

def sendMessage(message, messageType="log"):
    messageDict = {}
    if messageType == "log":
        messageDict["type"] = "log"
        messageDict["message"] = message
        emit("log", json.dumps(messageDict))
        socketio.sleep(0)
    elif messageType == "prog":
        messageDict["type"] = "prog"
        messageDict["message"] = message
        emit("prog", json.dumps(messageDict))
        socketio.sleep(0)
    elif messageType == "progMsg":
        messageDict["type"] = "progMsg"
        messageDict["message"] = message
        emit("progMsg", json.dumps(messageDict))
        socketio.sleep(0)
    elif messageType == "requireInput":
        messageDict["type"] = "requireInput"
        messageDict["message"] = message
        emit("requireInput", json.dumps(messageDict))
        socketio.sleep(0)
    else:
        emit("log", json.dumps("Failed Python message"))
        socketio.sleep(0)

@socketio.on('createFormattedExtract')
def createFormattedExtract(msg):
    emit("message", msg)
    commands = json.loads(msg)
    sendMessage(commands)
    Report = ER.ElentraReport(options=commands, reportType="Formatted Extract")
    Report.setExtractData(path=commands["extractDataFilePath"])
    Report.setLookupWorkbook(path=commands["lookupTableFilePath"])
    Report.setLookupTables()
    Report.createFormattedExtract()
    sendMessage("Your extract has been formatted!", "progMsg")
    sendMessage("100", "prog")
    del Report

@socketio.on('createGeneratedReport')
def createGeneratedReport(msg):
    emit("message", msg)
    commands = json.loads(msg)
    sendMessage(commands)
    Report = ER.ElentraReport(options=commands, reportType="Full Report")
    Report.setExtractData(path=commands["extractDataFilePath"])
    Report.setLookupWorkbook(path=commands["lookupTableFilePath"])
    Report.setLookupTables()
    Report.createFormattedExtract()
    Report.createResidentAnalysis()
    sendMessage("Your report has been generated!", "progMsg")
    sendMessage("100", "prog")
    del Report

@socketio.on('message')
def handleMessage(msg):
    send(msg, broadcast=True)


def shutdown_server():
    func = request.environ.get('werkzeug.server.shutdown')
    if func is None:
        raise RuntimeError('Not running with the Werkzeug Server')
    func()

@app.route('/')
def hello_world():
    return "Hello world"

@app.route('/shutdown')
def shutdown():
    shutdown_server()
    return 'Server shutting down...'

if __name__ == '__main__':
    socketio.run(app, debug=True)