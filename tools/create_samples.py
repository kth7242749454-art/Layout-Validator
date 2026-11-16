from pathlib import Path
import fitz
from openpyxl import Workbook

def create_samples(folder):
    folder=Path(folder); folder.mkdir(parents=True,exist_ok=True)
    book=Workbook(); sheet=book.active
    sheet.append(['Old_Instrument_Name','Instrument_Name','JUNCTION_BOX'])
    sheet.append(['9001-FT-102','9001-FT-101','9001-DJB-101'])
    sheet.append([None,'9001-FT-104','9001-DJB-102'])
    book.save(folder/'instrument_index.xlsx'); book.close()
    document=fitz.open(); page=document.new_page(width=700,height=500)
    page.insert_text((40,40),'SYNTHETIC EXAMPLE - NOT A PROJECT DRAWING',fontsize=14)
    page.insert_text((40,75),'INSTRUMENT LIST',fontsize=12)
    for x,label in [(40,'NO'),(95,'TAG NUMBER'),(300,'ELEVATION'),(430,'J/B NO.')]:
        page.insert_text((x,110),label,fontsize=10)
    rows=[('9001-FT-101','9001-DJB-101'),('9001-FT-102','9001-DJB-101'),('9001-FT-103','9001-DJB-103'),('9001-FT-104','9001-DJB-999')]
    for i,(tag,jb) in enumerate(rows,1):
        y=110+i*30
        for x,label in [(40,str(i)),(95,tag),(300,'100.00'),(430,jb)]: page.insert_text((x,y),label,fontsize=10)
    document.set_metadata({'title':'Synthetic Layout Validation Example','author':'Synthetic example generator'})
    document.save(folder/'synthetic_layout.pdf'); document.close()
    return folder

if __name__=='__main__': print('Created synthetic examples in',create_samples(Path(__file__).resolve().parents[1]/'samples'))
