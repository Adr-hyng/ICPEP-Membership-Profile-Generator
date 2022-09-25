import os
from re import search as regexSearch
from docxtpl import DocxTemplate
import pandas as pd
from datetime import datetime
from docx2pdf import convert

EXCEL_PATH = "Datasheet.xlsm"

os.chdir(os.path.dirname(os.path.abspath(__file__)))

class User:
    
    def __init__(self, name):
        self.name = [li.strip() for li in name]
        self.first_name = self.name[1]
        self.last_name = self.name[0]
        self.middle_init = self.name[2] if len(self.name) > 2 else ""



class MSWord:
    def __init__(self, file_name) -> None:
        self.user = None
        self.convertPDF = None
        self.file_name = file_name
        self.df = pd.read_excel(EXCEL_PATH, sheet_name = "Data")
        
        
    def start(self, convertPDF = False):
        self.convertPDF = convertPDF
        for self.record in self.df.to_dict(orient = "records"):
            self.doc = DocxTemplate(f"{self.file_name}.docx")
            self.fill_info()
            
    def fill_info(self):
        self.fill_personal()
        self.insert_images()
        self.fill_const()
        self.fill_contact_number()
        self.fill_date_entered()
        self.fill_organization_position()
        self.fill_talents_skills()
        self.fill_other_organization()
        self.valid_format()
        self.fill_docx()
        
            
    def insert_images(self):
        # Put images to generator
        # Link: https://signaturely.com/online-signature/draw/
        if not pd.isnull(self.record["user_id"]): self.doc.replace_pic("user_profile.jpg", f"Personal\ID\{self.user.first_name} {self.user.last_name}.jpg")
        if not pd.isnull(self.record["signature"]): self.doc.replace_pic("user_signature.png", f"Personal\Signatures\{self.user.first_name} {self.user.last_name}.png")
        if not pd.isnull(self.record["president_signature"]): self.doc.replace_pic("no_president_sign.png", f"CONSTANT\president_sign.png")
            
    def fill_personal(self):
        self.user = User(self.record["name"].split(","))
        father = User(self.record["fathers_name"].split(","))
        mother = User(self.record["mothers_name"].split(","))
        
        self.insert_images()
        
        self.record["last_name"], self.record["first_name"], self.record["mid"]  = (self.user.last_name, self.user.first_name, self.user.middle_init)
        self.record["father_surname"], self.record["father_first"], self.record["father_mid"]  = (father.last_name, father.first_name, father.middle_init)
        self.record["mother_surname"], self.record["mother_first"], self.record["mother_mid"]  = (mother.last_name, mother.first_name, mother.middle_init)
        self.record["printedName"] = f"{self.user.first_name} {self.user.middle_init}. {self.user.last_name}".upper()
        
        
    def fill_const(self):
        if any((pd.isnull(self.record["president_name"]), pd.isnull(self.record["adviser_name"]), pd.isnull(self.record["assist_adviser_name"]))): return
        president = User(self.record["president_name"].split(","))
        adviser = User(self.record["adviser_name"].split(","))
        assist_adviser = User(self.record["assist_adviser_name"].split(","))
        
        self.record["president_name"] = f"{president.first_name} {president.middle_init}. {president.last_name}".upper() if not pd.isnull(self.record["president_name"]) else ""
        self.record["assist_adviser_name"] = f"Engr. {assist_adviser.first_name} {assist_adviser.middle_init}. {assist_adviser.last_name}".upper() if not pd.isnull(self.record["assist_adviser_name"]) else ""
        self.record["adviser_name"] = f"Engr. {adviser.first_name} {adviser.middle_init}. {adviser.last_name}".upper() if not pd.isnull(self.record["adviser_name"]) else ""
        
    def fill_contact_number(self):
        self.record["contact"] = f"0{self.record['contact']}"
        
    def fill_date_entered(self):
        self.record["date_entered"] = self.record["date_entered"].strftime("%B %d, %Y")
            
    
    def fill_organization_position(self):
        self.record["organization_position"] = "Member" if self.record["organization_position"] else ""
        
    def fill_talents_skills(self):
        talent = {i+1:skill for i, skill in enumerate([li.strip() for li in self.record["talents"].split(",")])}
        for key, val in talent.items():
            self.record[f"talent{key}"] = val
        
    def fill_other_organization(self):
        org_membership = {i+1:membership for i, membership in enumerate([li.strip() for li in self.record["other_organization_membership"].split(",")])}
        org_position = {i+1:position for i, position in enumerate([li.strip() for li in self.record["other_organization_position"].split(",")])}
        for key, val in org_membership.items():
            self.record[f"membership{key}"] = val
            self.record[f"position{key}"] = org_position.get(key)
            
    
    def valid_format(self):
        for k, v in self.record.items():
            if pd.isnull(self.record[k]):
                self.record[k] = ""
            
            if not isinstance(v, (int, float, datetime)) and not k in ["printedName", "assist_adviser_name", "adviser_name", "president_name", "email"] and not (regexSearch("^membership[0-9]+$", k)):
                self.record[k] = v.title()
        
    def fill_docx(self):
        self.doc.render(self.record)
        output_path = f"Output/{self.user.first_name}-{self.user.last_name}_Membership_Profile.docx"
        self.doc.save(output_path)
        if self.convertPDF:
            convert(output_path, f"{output_path.replace('.docx', '.pdf')}", True)
            os.remove(output_path)

def main():
    _ms_word = MSWord(file_name = "Membership")
    _ms_word.start(convertPDF=True)
        
        
if __name__ == "__main__":
    main()
