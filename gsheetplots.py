import argparse
import sys
import glob
from jinja2 import Environment, FileSystemLoader
import os
import os.path
import numpy as np
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

try:
    AUTHKEY = os.environ['GSHEETAUTHKEY']
except KeyError:
    sys.exit(1)

def handleSpreadsheet(spreadsheetID):
    service = build('sheets', 'v4', developerKey=AUTHKEY)

    # Call the Sheets API
    sheet = service.spreadsheets()

    namedRanges = sheet.get(spreadsheetId=spreadsheetID).execute()['namedRanges']

    ranges = {}

    for r in namedRanges:
        if r['name'].startswith('wiki_'):
            rangeName = r['name']
            data = sheet.values().get(spreadsheetId=spreadsheetID, range=rangeName).execute()
            values = data.get('values', [])
            
            if not values:
                return None
            else:
                values = np.asarray(values)
            categories = values[0,1:].tolist()
            series = []
            for item in values[1:]:
                itm = {}
                itm['name'] = item[0].tolist()
                itm['data'] = item[1:].tolist()
                series.append(itm)
            ranges[rangeName] = {'series':series,'categories':categories}
    return ranges

def dir_path(path):
    if os.path.isdir(path):
        return path
    else:
        raise argparse.ArgumentTypeError(f"readable_dir:{path} is not a valid path")

def getTemplateFolder(rootpath, path):
    while path != os.path.abspath(os.path.join(rootpath)):
        tdir = os.path.join(path, "templates")
        if os.path.isdir(tdir):
            return tdir
        path = os.path.abspath(os.path.join(path, os.pardir))
    return os.path.join(rootpath, "templates")

def renderToFile(templateDir,template_name,outdir,ranges):
    env = Environment(loader=FileSystemLoader(templateDir))
    template = env.get_template(template_name)
    for key in ranges:
        output_from_parsed_template = template.render(categories=ranges[key]['categories'],series=ranges[key]['series'])
        fname = "{}_{}.txt".format(key,os.path.splitext(template_name)[0])
        with open(os.path.abspath(os.path.join(outdir, fname)), "w") as file:
            file.write(output_from_parsed_template)
        subRange = {}
        for index, cat in enumerate(ranges[key]['categories']):
            subRange[key] = {}
            subRange[key]['categories'] = cat
            subRange[key]['series'] = []
            for s in ranges[key]['series']:
                d = []
                d.append(s['data'][index])
                subRange[key]['series'].append({'name':s['name'],'data':d}) 
            
            fname = "{}_{}_{}.txt".format(key,cat.lower(),os.path.splitext(template_name)[0])
            output_from_parsed_template = template.render(categories=subRange[key]['categories'],series=subRange[key]['series'])
            with open(os.path.abspath(os.path.join(outdir, fname)), "w") as file:
                file.write(output_from_parsed_template)


def main():
    parser = argparse.ArgumentParser(description='Fill Templates with data from G-Sheets')
    parser.add_argument('rootDir', type=dir_path)
    args = parser.parse_args()
    rootDir = os.path.abspath(args.rootDir)
    for sheetfile in  glob.glob(rootDir + '/*/gsheetid.txt'):
        sheetid = ""
        with open(sheetfile, 'r') as file:
            sheetid = file.read().replace('\n', '')
        if 44 != len(sheetid):
            print("String {} doesn't look like a SheetId".format(childDir))
            continue
        sheetDir = os.path.split(sheetfile)[0]
    
        ranges  = handleSpreadsheet(sheetid)
        childDir = os.path.abspath(os.path.join(rootDir, sheetDir))
        templatedir = getTemplateFolder(rootDir, childDir)
        for template in glob.glob(os.path.join(childDir,'templates/*.txt')):
            template = os.path.basename(template)
            renderToFile(templatedir,template,childDir,ranges)


if __name__ == '__main__':
    main()
