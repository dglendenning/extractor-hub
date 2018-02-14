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
    url = ("***REMOVED***"
           "***REMOVED***")

    r = get(url)
    v = r.headers["content-disposition"]
    v = v[v.find("Extractor Hub v")+14:v.find(".exe")]

    if (v != __version__ and sorted(v, __version__)[0] != __version__):
        path = os.join(os.getcwd(), "Extractor Hub")
        with open(path, "wb") as file:
            # get request
            # write to file
            file.write(r.content)
        os.execv(path, [''])

    extractors.main()




















if __name__ == "__main__":
    main()
