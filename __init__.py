"""Auto-update if there are updates, otherwise launch main script."""
import sys
import os
from requests import get
import mal_data as mal
import extractors


if getattr(sys, 'frozen', False):
    # running in a bundle
    _mei_dir = sys._MEIPASS
else:
    # running live
    _mei_dir = os.path.split(__file__)

_version_file = os.path.join(_mei_dir, "version.txt")

with open(_version_file, 'r') as file:
    __version__ = file.readline()

_ver_url = mal.update_ver_url
_app_url = mal.update_app_url
_app_name = "Extractor Hub.exe"
_app_path = os.path.join(os.getcwd(), _app_name)


def download(url, file_name):
    """Use get() to download and save a file from the web."""
    # open in binary mode
    with open(file_name, "wb") as file:
        # get request
        response = get(url)
        # write to file
        file.write(response.content)


def update():
    """Get file from online and overrwrite this app with it.

    Only called once we know there's an update, since it closes the
    program.
    """
    global _app_path
    global _app_url

    try:
        os.rename(_app_path, "OLD.deleteme")
        download(_app_url, _app_path)
        restart_program()
    except Exception as ex:
        raise ex


def restart_program():
    """Restarts the current program.

    Note that this function does not return. Any cleanup action, like
    saving data, must be done before calling this function.
    """
    os.execv(sys.executable, ['sudo python'] + sys.argv)


def version_from_header(string):
    """Extract version number from htpps header."""
    head = string.find('filename="') + len('filename="')
    tail = string.find('.txt')
    return string[head:tail]


def update_available():
    """Return true if local version is older than latest version."""
    global _ver_url
    global __version__
    # Get most recent version.
    header = get(_ver_url).headers["content-disposition"]
    web_version = version_from_header(header)
    # Convert version numbers to lists to compare
    web_vlist = [int(item) for item in web_version.split('.')]
    loc_vlist = [int(item) for item in __version__.split('.')]
    for web, local in zip(web_vlist, loc_vlist):
        if web > local:     # There's a new version.
            return True
        elif local > web:   # We're running a dev build.
            return False
        return False    # We're up to date.


def main():
    """Check for updates, then either update or launch."""
    # If we have just updated, remove old version.
    if os.path.exists("OLD.deleteme"):
        os.remove("OLD.deleteme")
    # If there's an update, install it.
    if update_available():
        update()
    # Otherwise, launch the app.
    else:
        extractors.main(__version__)


if __name__ == "__main__":
    main()
