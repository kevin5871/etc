# Code From : https://sqa.stackexchange.com/questions/41928/how-to-autoupdate-chrome-driver-in-selenium 
# For : downloading driver
# Code From : https://stackoverflow.com/questions/57441421/how-can-i-get-chrome-browser-version-running-now-with-python 
# For : getting chrome version

import requests
from win32com.client import Dispatch
import zipfile
import tempfile
import os

td = tempfile.TemporaryDirectory()
path = td.name
def get_version_via_com(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    try:
        version = parser.GetFileVersion(filename)
    except Exception:
        return None
    return version

paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
         r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
version = list(filter(None, [get_version_via_com(p) for p in paths]))[0]
print("Current Version : " + version)
ver = version.split('.')
version = ver[0]

url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_'
url_file = 'https://chromedriver.storage.googleapis.com/'
file_name = 'chromedriver_win32.zip'

#version = input()
version_response = requests.get(url + version)
if version_response.text:
    file = requests.get(url_file + version_response.text + '/' + file_name)
    with open(os.path.join(path, file_name), "wb") as code:
        code.write(file.content)

zipfile.ZipFile(os.path.join(path, "chromedriver_win32.zip"), "r").extractall()
td.cleanup()
