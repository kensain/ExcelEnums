I parsed all of the Excel Enumerations from the microsoft website. I initially wanted to turn it into an Autohotkey object for easier scripting, but right now I feel kinda lazy, so I decided to leave the json file as is. See **excel_enums.json**. A little excerpt (see Microsoft [webpage](https://learn.microsoft.com/en-us/office/vba/api/excel.xlconditionvaluetypes "https://learn.microsoft.com/en-us/office/vba/api/excel.xlconditionvaluetypes") for this enumeration):

```JSON
"xlconditionvaluetypes": {
        "xlConditionValueAutomaticMax": {
            "Name": "xlConditionValueAutomaticMax",
            "Value": 7,
            "Description": "The longest data bar is proportional to the maximum value in the range."
        },
        "xlConditionValueAutomaticMin": {
            "Name": "xlConditionValueAutomaticMin",
            "Value": 6,
            "Description": "The shortest data bar is proportional to the minimum value in the range."
        },
        "xlConditionValueFormula": {
            "Name": "xlConditionValueFormula",
            "Value": 4,
            "Description": "Formula is used."
        },
        "xlConditionValueHighestValue": {
            "Name": "xlConditionValueHighestValue",
            "Value": 2,
            "Description": "Highest value from the list of values."
        },
        "xlConditionValueLowestValue": {
            "Name": "xlConditionValueLowestValue",
            "Value": 1,
            "Description": "Lowest value from the list of values."
        },
        "xlConditionValueNone": {
            "Name": "xlConditionValueNone",
            "Value": "-1",
            "Description": "No conditional value."
        },
        "xlConditionValueNumber": {
            "Name": "xlConditionValueNumber",
            "Value": 0,
            "Description": "Number is used."
        },
        "xlConditionValuePercent": {
            "Name": "xlConditionValuePercent",
            "Value": 3,
            "Description": "Percentage is used."
        },
        "xlConditionValuePercentile": {
            "Name": "xlConditionValuePercentile",
            "Value": 5,
            "Description": "Percentile is used."
        }
    },
```

For some reason negative numbers didn't convert to the type int, so they are stored a string. I don't know why. I provided a python script I used, in case if anyone is interested. I suggest setting it up via `venv`, `requirements.txt` is provided.
