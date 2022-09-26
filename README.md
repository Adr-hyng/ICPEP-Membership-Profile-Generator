# ICpEP Membership Profile Generator

![image](https://user-images.githubusercontent.com/95139246/192171805-768b9665-91d7-46cf-847c-7f0a4ed7c939.png)


#

## Format:
```
General:
- Family Info

Leave it blank <-optinal> or None <-required>

- I. Name / Father's Name / Mother's Name
- @params
- lastName, firstName <[-optional] secondName> <[-optional] suffix>, middleInitial | maidenName (mother) <[-optional]> 
- Ex.
- Dela Cruz, Juan Bantoy Jr., A
- Dela Cruz, Juan Sr., A

- II. Current Position in this Organization (Leave it blank if Member) <-optional>
- @default: Member
- Ex.
- <empty> = Member
- Treasurer = Treasurer

- III. Talents / Skills <-optional>
- Commas as separator, and only 4 talents can only be putted.
- @params
- Ex.
- Dancing, Singing, Communication Skills
- Dancing

- IV. Other Orgs <-optional>
- Commas as separator, and only 3 orgs can only be putted.
- @params
- Ex.
- USC, DELTA, IECEP
- USC
```

#

## Creator Note:
- Follow the Format (Above this)
- 1 Image / Picture per Upload File of User.

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

![image](https://user-images.githubusercontent.com/95139246/192169490-82efcc3a-4ff1-43a3-a8da-e92ee850c2ae.png)


#### User:
- Log-on Google Account
- Fillout Membership Profile Google Form
- Put Personal Information in Google Form
- Upload Images and Signature that came from Signaturely. (Make sure width is under W and i, like shown below)

![image](https://user-images.githubusercontent.com/95139246/192169857-eec48bb6-22b1-4617-9d8b-874882d4b845.png)

- Ignore Others and Submit.
- Done

#### Admin:
- Create Google Form similar questions and its sequence to [THIS](https://docs.google.com/forms/d/1s8AaJSy2avSU80pV6uH1AdfvGnhmsL7igPiskKmrurk/edit#responses) template Google Form.
- In Google Form, click right side of Send Button that 3-dots, then there you will see "Script Editor", click it.
Paste this Code there. [APP SCRIPT SOURCE CODE](https://paste.pythondiscord.com/cixisonone), after pasting, save and tap the Alarm Clock Icon in Left Side, which is the trigger icon. Add new Trigger, then just change the last to "On Submit". Now, you are all set.

![image](https://user-images.githubusercontent.com/95139246/192169459-6c471b67-dca8-49dd-9b20-28861073d1e9.png)

![image](https://user-images.githubusercontent.com/95139246/192169434-a6e01c24-e000-425e-906e-60156e551fcb.png)


- Make sure to change the Google Form ID and Google Drive ID to the code you've pasted.
```js
const GOOGLE_FORM_ID = "https://docs.google.com/forms/d/<ID HERE>";
const PROFILE_DRIVE_ID = "https://drive.google.com/drive/folders/<ID HERE>";
const SIGNATURES_DRIVE_ID = "https://drive.google.com/drive/folders/<ID HERE>";
```
- After following the instructions, you are good to go, every submitted responses' uploaded file will be named according to their response answers of name or the first Question in Google Form.
- Just Download the File and Move the Files accordingly.
- Get Responses through Google Spreadsheets
- Get Responses through Cell in Spreadsheet and Copy it not including the Timestamp.
- Paste it to "Datasheet.xlsm" Data Sheet starting from name.
- After Pasting required personal info in cell of MS Excel, save it. 
- Run the python Script.


## Developer's Workspace

### Testing:
- It needs to be tried with at least 10 filled membership profile if it works.

### Possible Feature:
- Automatically Move Files that has "* - ID.jpg" namings to `Personal\ID\here` and "* - Signature.png" to `Personal\Signatures\here`.
- Optimize Error Handling
- Add Regular Expression in Google Form such: **Talents, Age, Membership to other orgs, Position to other orgs, Contact Number, Email Address, Father's Name, Mother's Name, Name.**

## How to clone?

```js
git clone https://github.com/Adr-hyng/ICPEP-Membership-Profile-Generator.git
```

then,

run this on CLI.


```py
pip install -r "/path/to/requirements.txt"
```

After installing you can now run the script in your machine.

