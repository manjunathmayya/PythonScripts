# -*- coding: utf-8 -*-
"""
This script opens and displays the links in Links.csv file in form of html.
After running this script, type localhost:5000 in web browser, links will be displayed
Also typing ipaddress:5000 in browser will also display the links.
"""
import pandas as pd
from flask import Flask, render_template

app = Flask(__name__)
        
@app.route("/")
def main():
    
    pd.set_option('display.max_colwidth', -1)
    pd.set_option('colheader_justify', 'center')
     
    data = pd.read_csv('Links.csv')
    data['Links'] = data['Links'].apply(lambda x: '<a href="{}">{}</a>'.format(x,x))    
   
    data.to_html('templates/mytemplate.html',escape=False)
    return render_template('mytemplate.html')

if __name__ == "__main__":
    app.run(host='0.0.0.0', debug=True, port=5000)
