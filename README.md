# ICpEP Membership Profile Generator

## Developer Test
- It needs to be tried with at least 10 filled membership profile if it works.

# 
## How does it work?
### Requirements:
- MS Excel
- [Google Spreadsheet (Responses of Google Forms)](https://docs.google.com/spreadsheets/d/1XmEEnh9Wd3oCinoYBo0pwT3PbWjUFvGjiHf0hEcZg9U/edit?resourcekey&usp=forms_web_b#gid=1708594941)
- [Google Forms](https://docs.google.com/forms/d/1s8AaJSy2avSU80pV6uH1AdfvGnhmsL7igPiskKmrurk/edit)
- MS Word
- [Signaturely](https://signaturely.com/online-signature/draw/)

### - Using MS Excel It can get it's input and transfer It to MS Docs to be converted by substituting the placeholders.
### - Use Google Form to easily gather data or personal information to be generated.
### - Use Google Spreadsheet by only copying cells to be pasted in MS Excel
### - Using Signaturely for consistent e-signatures to be generated since all layout and width of brush is customizable.
#

## How to use?

### Steps:
```js
Directory:

|/Root
├── CONSTANT
├── Output
├── Personal
├──── ID
├────── *.jpg
├──── Signatures
├────── *.png
├── Membership.docx
├── main.py
├── Datasheet.xlsm

5 directories, 3 files not including pictures.
```

#### User:
- Log-on Google Account
- Fillout Membership Profile Google Form
- Put Personal Information in Google Form
- Uplaod Images and Signature that came from Signaturely.
- Ignore Others and Submit.
- Wait...

### Admin:
- Create Google Form similar questions and its sequence to [THIS](https://docs.google.com/forms/d/1s8AaJSy2avSU80pV6uH1AdfvGnhmsL7igPiskKmrurk/edit#responses) template Google Form.
- Wait for Responses
- Get Responses through Google Spreadsheets
- Get Responses through Cell in Spreadsheet and Copy it not including the Timestamp.
- Paste it to "Datasheet.xlsm" Data Sheet starting from name.
- After Pasting required personal info in cell of MS Excel, save it. 
- Download Pictures such as User Profile and User Signature
- Move User Profile to /Root/Personal/ID/ and Move User Signature to /Root/Personal/Signatures/.
- Run Python File "main.py"
- See output in Output Folder or Directory.


## How to clone?

```js
git clone <>
```

then,

run this on CLI.


```py
pip install -r "/path/to/requirements.txt"
```

After installing you can now run the script in your machine.

