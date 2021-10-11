
# Link counter

Script allows to add links to .xlsx file to track your activity.




## Features

- Add current link from browser
- Extract ID from it (last digits)
- Divide links by date adding a separate sheet for a day

  
## Instructions:

    Add ticket - Alt + S
    Get ticket count - Alt + C
    Open sheet copy - Alt + O
  
## Optional

It is possible to set the links that script can accept. Edit the following variable in contstant.py, example:

```
LINK_STARTSwITH = 'https://'
```

  
## Installation

At least Python 3.6 is required (f-string)

Create venv

```bash
  pip install -r requirements.txt
```
And then run script:

```bash
  python main.py
```
    