#!/usr/bin/python3

import os
import flask
import httplib2
from apiclient import discovery
from apiclient.http import MediaIoBaseDownload, MediaFileUpload
from oauth2client import client
from oauth2client import tools
from googleapiclient.discovery import build
from oauth2client.file import Storage
from flask_wtf import FlaskForm
from wtforms import TextField
from requests import get
from bs4 import BeautifulSoup

app = flask.Flask(__name__)
app.config.from_object(__name__)
app.config['SECRET_KEY'] = '7d441f27d441f27567d441f2b6176a'

@app.route('/')
def index():
    credentials = get_credentials()
    if credentials == False:
        return flask.redirect(flask.url_for('oauth2callback'))
    elif credentials.access_token_expired:
        return flask.redirect(flask.url_for('oauth2callback'))
    else:
        print('now calling fetch')
        form = SlideForm()
        return flask.render_template("index.html", form=form)

@app.route('/create-slides', methods=['POST'])
def createslides():
    creds = get_credentials()
    form = SlideForm()
    information = get_info(form.name.data)
    paragraphs = information["paragraphs"]
    title = information["title"]
    PBF = {
        "propertyState": "RENDERED",
        "solidFill": {
            "color": {
                "rgbColor": {
                    "red": 1.0,
                    "green": 0.0,
                    "blue": 0.0,
                }
            },
            "alpha" : 1.0
        }
    }
    # Create presentation
    body = {
        'title': title,
    }
    slides_service = build('slides', 'v1', credentials=creds)
    presentation = slides_service.presentations() \
    .create(body=body).execute()
    print('Created presentation with ID: {0}'.format(
        presentation.get('presentationId')))

    # populates the title
    first_slide = presentation.get('slides')[0]
    header_text = first_slide.get('pageElements')[0]

    pt350 = {
        'magnitude': 350,
        'unit': 'PT'
    }
    requests = []
    requests.append(
        {
            'insertText': {
                'objectId': header_text.get("objectId"),
                'insertionIndex': 0,
                'text': title,
            }
        })
    element_id = 0
    page_id = 0
    for p in paragraphs:
        new_slide = {
            'createSlide': {
                'objectId': "hacker" + str(page_id),
                'insertionIndex': '1',
                'slideLayoutReference': {
                    'predefinedLayout': 'BLANK'
                },
            },
        }
        requests.append(new_slide)
        text_box = {
            'createShape': {
                'objectId': "MY-TEXT" + str(element_id),
                'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': "hacker" + str(page_id),
                    'size': {
                        'height': pt350,
                        'width': pt350
                    },
                    'transform': {
                        'scaleX': 1,
                        'scaleY': 1,
                        'translateX': 350,
                        'translateY': 100,
                        'unit': 'PT'
                    }
                }
            }
        }
        requests.append(text_box)
        text = {
            # Insert text into the box, using the supplied element ID.
            'insertText': {
                'objectId': "MY-TEXT" + str(element_id),
                'insertionIndex': 0,
                'text': p,
            }
        }
        requests.append(text)
        color = {
            "updatePageProperties": {
                "objectId": "hacker" + str(page_id),
                "pageProperties": {
                "pageBackgroundFill": PBF
                },
                "fields": "pageBackgroundFill"
            }
        }
        requests.append(color)
        page_id = page_id + 1
        element_id = element_id + 1
    body = {
        'requests': requests
    }
    response = slides_service.presentations() \
    .batchUpdate(presentationId= presentation.get('presentationId'), body=body).execute()

    display_url = "https://docs.google.com/presentation/d/" + presentation.get('presentationId')

    return flask.render_template("index.html", form=form, url = display_url)


class SlideForm(FlaskForm):
   name = TextField("Your Text")

def get_info(url):
    results = {}
    response = get(url)
    html_soup = BeautifulSoup(response.text, 'html.parser')
    paragraphs = [para.getText() for para in html_soup.find_all('p')]
    results["paragraphs"] = paragraphs
    title = html_soup.find('h1').getText()
    results["title"] = title
    return results


@app.route('/oauth2callback')
def oauth2callback():
    flow = client.flow_from_clientsecrets('client_id.json',
	        scope='https://www.googleapis.com/auth/drive',
            redirect_uri=flask.url_for('oauth2callback', _external=True)) # access drive api using developer credentials
    flow.params['include_granted_scopes'] = 'true'
    if 'code' not in flask.request.args:
        auth_uri = flow.step1_get_authorize_url()
        return flask.redirect(auth_uri)
    else:
        auth_code = flask.request.args.get('code')
        credentials = flow.step2_exchange(auth_code)
        open('credentials.json','w').write(credentials.to_json()) # write access token to credentials.json locally 
        return flask.redirect(flask.url_for('index'))

def get_credentials():
	credential_path = 'credentials.json'

	store = Storage(credential_path)
	credentials = store.get()
	if not credentials or credentials.invalid:
		print("Credentials not found.")
		return False
	else:
		print("Credentials fetched successfully.")
		return credentials

def fetch(query, sort='modifiedTime desc'):
	credentials = get_credentials()
	http = credentials.authorize(httplib2.Http())
	service = discovery.build('drive', 'v3', http=http)
	results = service.files().list(
		q=query,orderBy=sort,pageSize=10,fields="nextPageToken, files(id, name)").execute()
	items = results.get('files', [])
	return items

def download_file(file_id, output_file):
	credentials = get_credentials()
	http = credentials.authorize(httplib2.Http())
	service = discovery.build('drive', 'v3', http=http)
	#file_id = '0BwwA4oUTeiV1UVNwOHItT0xfa2M'
	request = service.files().export_media(fileId=file_id,mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
	#request = service.files().get_media(fileId=file_id)
	
	fh = open(output_file,'wb') #io.BytesIO()
	downloader = MediaIoBaseDownload(fh, request)
	done = False
	while done is False:
		status, done = downloader.next_chunk()
		#print ("Download %d%%." % int(status.progress() * 100))
	fh.close()
	#return fh
	
def update_file(file_id, local_file):
	credentials = get_credentials()
	http = credentials.authorize(httplib2.Http())
	service = discovery.build('drive', 'v3', http=http)
	# First retrieve the file from the API.
	file = service.files().get(fileId=file_id).execute()
	# File's new content.
	media_body = MediaFileUpload(local_file, resumable=True)
	# Send the request to the API.
	updated_file = service.files().update(
		fileId=file_id,
		#body=file,
		#newRevision=True,
		media_body=media_body).execute()
		
if __name__ == '__main__':
	if os.path.exists('client_id.json') == False:
		print('Client secrets file (client_id.json) not found in the app path.')
		exit()
	import uuid
	app.secret_key = str(uuid.uuid4())
	app.run(debug=True)