"""Parser module to convert LBP pdfs into csv files.

Note:
    pypdf doesn't work well for LBP documents

Usage:
  pattern            type documents (par ex: *.pdf)

options:
  -h, --help         show this help message and exit
  -i, --input INPUT  input directory
  -d, --dest DEST    output directory

Example:
    py pdf-extract *.pdf

"""

import argparse
import glob
from glob import glob
import pymupdf
import tabula
import pandas
import re
from pathlib import Path
from os import makedirs

output_dir = r'./'

def rebuild(df):
   dfb=rename_cols(drop_empty_col(df))
   dfb["Débit"] = dfb["Débit"].str.replace(' ','').str.replace(',','.').astype(float)
   dfb["Crédit"] = dfb["Crédit"].str.replace(' ','').str.replace(',','.').astype(float)
   return dfb.fillna('')

def drop_empty_col(df):
   list_unnamed=[(i) for i in df.columns if i.startswith('Unnamed')]
   return df.drop(columns=list_unnamed)

def rename_cols_0(df):
   return df.rename(columns={"Débit (¤)": "Débit",  "Crédit (¤)": "Crédit"})

def clean_operation(name):
    return re.sub(r"\d\d\/\d\d\s", "", name).strip()

def rename_cols(df):
   dfb = rename_cols_0(df)
   col0 = df.columns[0]
   if col0.endswith("Opération"):
       print(col0)
       dfb.insert(0, "Date", "")
       re = r'(\d\d\/\d\d)\s'
       dfb["Date"] = dfb[col0].str.extract(re)
       dfb[col0] = dfb[col0].apply(clean_operation)
       return dfb.rename(columns={"Date Opération": "Opération"})

   return dfb

#        y0  x0 y1  x1 for tabula
area = [[550,10,780,600], # first page
        [150,10,780,600], # inter page
        [150,10,420,600]] # last page

# Rect(x0,y0,x1,y1) for pymupdf

def extract_tables(filename):
    doc = pymupdf.open(filename)
    pageNb = 0
    combined_df = pandas.DataFrame()
    fin_tables = False
    for page in doc:

       pageNb+=1
       clip=[]
       rect_debut = page.search_for("Débit (¤)")
       rect_fin = page.search_for("Nouveau solde")
       # if rect_fin and pageNb > 1:
           # fin_tables = True 
       
       print("page",pageNb)
       if not rect_debut and fin_tables is False:
          exit(f"Erreur! rec-debut(Débit) pas trouvé; page {pageNb}")

       point_deb = rect_debut[0].tl # point at top left
       if rect_fin:
          point_fin = rect_fin[0].br # point at top left

       if not rect_fin:   # inter page
          clip = list(area[1])
          #clip[0] = point_deb.y
          
       elif pageNb==1:  # first page
          clip = list(area[0])
          #clip[0] = point_deb.y

       else:    # last page
          clip = list(area[2])
          clip[2] = point_fin.y
          fin_tables = True 

       df = tabula.read_pdf(filename, pages=pageNb, encoding='utf-8', area=[clip], multiple_tables=False)[0]
       if pageNb == 1:
           df.iloc[[0],[1]] = "Ancien solde"
       rebuild_df = rebuild(df)
       
       combined_df = pandas.concat([combined_df, rebuild_df], ignore_index=True, sort=False)
       print(rebuild_df.head())
       print(clip)
#       exit()

       if fin_tables:
          break
       
    return combined_df
       
def export_to_csv(filename):
   global output_dir
   
   df = extract_tables(filename)
   name = Path(filename).stem
   #output_dir = Path(filename).parent
   outFile = f"{output_dir}/{name}.csv"

   df.to_csv(outFile, index=False, sep=";", encoding="utf-8")

   outFile = f"{output_dir}/{name}.xlsx"
   writer = pandas.ExcelWriter(outFile,
                engine='xlsxwriter',
                engine_kwargs={'options': {'strings_to_numbers': True}})

   df.to_excel(writer, sheet_name="bank", index=False)

   writer.close()
   print("File " + outFile + " is created")




def main():
    global output_dir
    
    parser = argparse.ArgumentParser(description='Process pdf files')
    parser.add_argument("pattern", help="type documents (par ex: *.pdf)")
    parser.add_argument('-i','--input', help='input directory', default='./')
    parser.add_argument('-d','--dest', help='output directory', default='./')

    args = parser.parse_args()
    
    output_dir = args.dest
    
    makedirs(output_dir, exist_ok=True)

    print(args)

    pdf_files = []

    pdf_files.extend(f for f in Path(args.input).glob(args.pattern))
    
    for f in pdf_files:
        if not Path(f).name.endswith('.pdf'):
            pdf_files.remove(f)

    if not pdf_files:
        exit('Aucun fichier PDF trouvé')

    for file in pdf_files:
        print(file)
        export_to_csv(file)


if __name__ == "__main__":
    main()
