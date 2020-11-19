from flask import *  
from app import *
app = Flask(__name__)  
import final
import font_name
import os
from werkzeug.utils import secure_filename
import glob
import csv 

#Loading Homepage 
@app.route('/')
@app.route('/upload')   
def upload():  
    return render_template("upload.html")  

#Uploade File
@app.route('/success', methods = ['GET','POST'])  
def success():
    if request.method =='POST':
        f=request.files['file']
        file = str(f.filename)
        f.save(f.filename)
        #org_file[0] == pdf file and [1] == docx file
        org_file = final.file_convert(file)
        person_info = final.person_details(org_file[0])
        linkin_id = final.linkin(org_file[1])
        no_of_lines = final.no_lines(org_file[0])
        no_of_char = final.no_char(org_file[0])
        font_names = font_name.fontname(org_file[0])
        fontsize = final.font_size(org_file[1])
        total_table = final.count_tables(org_file[1])
        total_img = final.count_img(org_file[0])
        no_of_lines = ' , '.join([str(elem) for elem in no_of_lines])
        no_of_char = ' , '.join([str(elem) for elem in no_of_char])
        font_names = ' , '.join([str(elem) for elem in font_names])
        fontsize = ' , '.join([str(elem) for elem in fontsize])
        for fi in glob.glob("*.pdf"):
            if(fi != org_file[0]):
                os.remove(fi)
        for fi in glob.glob("*.docx"):
            if(fi != org_file[1]):
                os.remove(fi)
        return render_template("result.html",name=person_info['name'],mob=person_info['mobile_number'],
                               mail=person_info['email'],linkin=linkin_id,pages=person_info['no_of_pages'],
                               lines=no_of_lines,char=no_of_char,font=font_names,f_size=fontsize,table=total_table,
                               image=total_img)

#Running Homefile
if __name__ == '__main__':
    app.run(debug = True)