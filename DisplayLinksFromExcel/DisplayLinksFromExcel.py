# -*- coding: utf-8 -*-
"""
Created on Wed Jan 30 00:27:34 2019

@author: ic014090
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
