from bs4 import BeautifulSoup
import requests
from urllib.parse import urljoin
import json


def extract_links() -> list:
    r = requests.get(
        "https://learn.microsoft.com/en-us/office/vba/api/excel(enumerations)"
    )
    soup = BeautifulSoup(r.content, "html.parser")
    res = []
    links = soup.find_all("a")  # Find all elements with the tag <a>
    for link in links:
        if not isinstance(link.string, str):
            continue
        if link.string.startswith("Xl"):
            res.append(
                urljoin(
                    "https://learn.microsoft.com/en-us/office/vba/api/",
                    link.get("href"),
                )
            )
    return res


def extract_table(link: str):
    r = requests.get(link)
    soup = BeautifulSoup(r.content, "html.parser")
    name: str = link.split(".")[-1]
    tables = soup.find_all("table")

    for table in tables:
        res = {name: {}}
        headers = []
        for header in table.thead.find_all("th"):
            headers.append(header.text)
        for row in table.tbody.find_all("tr"):
            res[name][row.find_all("td")[0].text] = {
                header: convert_to_int(row.find_all("td")[headers.index(header)].text)
                for header in headers
            }
        return res


def convert_to_int(text: str) -> str | int:
    return int(text) if text.isdigit() else text


res = {}
links = extract_links()
link_counter = 1
for link in links:
    print(f"Parsing {link_counter}/{len(links)}: {link!r}")
    if extract_table(link) is None:
        print(f"{link!r} was skipped. Probably not tables there.")
        link_counter += 1
        continue
    for key, value in extract_table(link).items():
        res[key] = value
    link_counter += 1

# Serializing json
json_object = json.dumps(res, indent=4)

# Writing to sample.json
with open("xlEnums.json", "w") as outfile:
    outfile.write(json_object)

print("Parsing complete!")
