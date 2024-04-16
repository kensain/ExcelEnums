I parsed all of the Excel Enumerations from the microsoft website. I initially wanted to turn it into an Autohotkey object for easier scripting, but right now I feel kinda lazy, so I decided to leave the json file as is. See **excel_enums.json**.

For some reason negative numbers didn't convert to the type int, so they are stored a string. I don't know why. I provided a python script I used, in case if anyone is interested. I suggest setting it up via `venv`, `requirements.txt` is provided.
