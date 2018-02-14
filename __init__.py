"""__init__.py
"""
import extractors
from requests import get  # to make GET request
import sys
import os

__version__ = extractors.__version__
VERSION_URL = ("***REMOVED***"
             "***REMOVED***")
APP_URL = ("***REMOVED***"
           "***REMOVED***")
APP_NAME = "Extractor Hub.exe"
FILE_NAME = os.path.join(os.getcwd(), APP_NAME)

def download(url, file_name):
    # open in binary mode
    with open(file_name, "wb") as file:
        # get request
        response = get(url)
        # write to file
        file.write(response.content)

def update():
    """Download file from ***REMOVED*** and overrwrite the local app with it.
    Only called once we know there's an update (if this runs every time we
    never get to run the program)."""
    global FILE_NAME
    global APP_URL

    try:
        os.rename(FILE_NAME, "OLD.deleteme")
        download(APP_URL, FILE_NAME)
        restart_program()
    except Exception as ex:
        #print("Uh oh")
        raise ex


def restart_program():
    """Restarts the current program.
    Note: this function does not return. Any cleanup action (like
    saving data) must be done before calling this function."""
    os.execv(sys.executable, ['sudo python'] + sys.argv)

def launch():
    """After checking for updates, we run the actual program! Call only if
    no update is available."""
    extractors.main(__version__)

def main():
    global VERSION_URL
    if os.path.exists("OLD.deleteme"):
        os.remove("OLD.deleteme")
    r = get(VERSION_URL)
    contentstring = str(r.content)
    contentstring = contentstring.strip("b'\\r\\n")
    #print(contentstring)
    string = r.headers["content-disposition"]
    start = string.find('filename="') + len('filename="')
    stop = string.find('.txt')

    result = string[start:stop]
    #print(result)
    web_version = [int(item) for item in result.split('.')]
    my_version = [int(item) for item in __version__.split('.')]
    for web, local in zip(web_version, my_version):
        if web > local:
            update()
            break
        elif local > web:
            #print("Note: current version is newer than official version."
            #      " This is unsupported.")
            break
    #else:
        #print("No Update Available")

    launch()

if __name__ == "__main__":
    main()
