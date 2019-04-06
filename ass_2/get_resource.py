import requests
import csv
import os
import xlrd
from flask import render_template, Flask,jsonify
from ass_2.auth import login_required
from flask_restful import reqparse
from mongoengine import connect,StringField, Document,FloatField,ListField,EmbeddedDocument,EmbeddedDocumentField

app = Flask(__name__)

class time(EmbeddedDocument):
    offence_time = StringField(required = True, primary_key=True)
    nb_of_incidents =FloatField(required = True)
    rt_of_p_mill_pop = FloatField(required = True)

class offence_type(EmbeddedDocument):
    type_name = StringField(required = True, primary_key=True)
    trend_24 =StringField(required = True)
    trend_64 =StringField(required = True)
    Rank_of_2016 =StringField()
    time_rate = ListField(EmbeddedDocumentField(time))

class offence_group(EmbeddedDocument):
    group_name = StringField(required = True, primary_key=True)
    type = ListField(EmbeddedDocumentField(offence_type))

class region(Document):
    region_name = StringField(required = True, primary_key=True)
    content = ListField(EmbeddedDocumentField(offence_group))


def data_import_code(ele):
    b =open('NSWcode.csv', 'r')
    a = csv.reader(b)
    l = []
    for line in a:
        if line and line[1] == ele:
            l.append(data_import_name(line[0]))
    b.close()
    if any(l):
        return True
    return False


def data_import_name(region_name):
    ele = region_name.replace(' ','')
    URL = f'http://www.bocsar.nsw.gov.au/Documents/RCS-Annual/{ele}LGA.xlsx'
    a = requests.get(URL)
    if a.status_code != requests.codes.ok:
        return False
    with open(f'{ele}.xlsx','wb') as file:
        file.write(a.content)
    data = xlrd.open_workbook(f'{ele}.xlsx')
    sheet1 = data.sheets()[0]
    content = []
    for m in range(7,69):
        if sheet1.cell(m,0).value:
            offence_process = []
            for i in range(m,69):
                time_process = []
                for j in range(2,11,2):
                    time_process.append(time(sheet1.cell(5,j).value,sheet1.cell(i,j).value,sheet1.cell(i,j+1).value))
                if sheet1.cell(i,1):
                    offence_process.append(offence_type(sheet1.cell(i,1).value,str(sheet1.cell(i,12).value),str(sheet1.cell(i,13).value),str(sheet1.cell(i,14).value),time_process))
                else:
                    offence_process.append(offence_type(sheet1.cell(i, 0).value, str(sheet1.cell(i, 12).value),
                                                        str(sheet1.cell(i, 13).value), str(sheet1.cell(i, 14).value),
                                                        time_process))
                if sheet1.cell(m+1,0).value:
                    break
            content.append(offence_group(sheet1.cell(m,0).value,offence_process))
    result = region(region_name,content)
    connect(region_name)
    result.save()
    os.remove(f'{ele}.xlsx')
    return True


@app.route('/log_in',methods = ['POST'])
def index():
    return render_template('log_in.html')

@app.route('/admin/import/lganame',methods =['POST'])
@login_required
def import_entry_name():
    parser = reqparse.RequestParser()
    parser.add_argument('lganame',type =str)
    args = parser.parse_args()
    name = args.get('lganame')
    result = data_import_name(name)
    pass


@app.route('/admin/import/postcode',methods =['POST'])
@login_required
def import_entry_code():
    parser = reqparse.RequestParser()
    parser.add_argument('postcode',type =str)
    args = parser.parse_args()
    code =args.get('postcode')
    result = data_import_code(code)


if __name__ == '__main__':
    connect(host = 'mongodb://admin:admin@ds063946.mlab.com:63946/ass_2')
    data_import_name('bluemountains')
