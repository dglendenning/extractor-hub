"""__init__.py
"""
import extractors
from requests import get  # to make GET request

__version__ = "v0.0.1"

#Check for updates
def download(url, file_name):
    # open in binary mode
    with open(file_name, "wb") as file:
        # get request
        response = get(url)
        # write to file
        file.write(response.content)


def main():
    print("Starting main function...")
    #url = ("***REMOVED***"
    #       "***REMOVED***")

    # URL of the VERSION NUMBER TEXT FILE, not the real app file.
    # This file has its NAME changed to each new version number upon release.
    url = ("***REMOVED***"
         "***REMOVED***")
    print("URL set!")
    r = get(url)
    print("Downloaded file")
    v = r.headers["content-disposition"]
    print("Content Disposition:", v)
    v = v[v.find("Extractor Hub v")+14:v.find(".exe")]

    if (v != __version__ and sorted(v, __version__)[0] != __version__):
        path = os.join(os.getcwd(), "Extractor Hub")
        with open(path, "wb") as file:
            # get request
            # write to file
            file.write(r.content)
        os.execv(path, [''])

    extractors.main()



















print("Running __init__")
if __name__ == "__main__":
    print("Name is __main__")
    main()
    print("End main")
