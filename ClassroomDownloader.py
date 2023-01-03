"""
* @ file ClassroomDownloader.py
* @ author Felix Kröhnert(felix.kroehnert@online.de)
* @ brief Download tool to download all files in the Classroom Stream
* @ version 0.1
* @ date 2023-01-02
*
* @ copyright Copyright(c) 2023 Felix Kröhnert
*
"""
#
# Requirements:
# - Google Cloud Application
# - Test Users that have access to the specified Classroom
# - Classroom API and Drive API
# - following scopes for OAuth:
# https://www.googleapis.com/auth/classroom.courseworkmaterials.readonly
# https://www.googleapis.com/auth/classroom.courses.readonly
# https://www.googleapis.com/auth/classroom.student-submissions.me.readonly
# https://www.googleapis.com/auth/classroom.announcements.readonly
# https://www.googleapis.com/auth/drive
#
# - OAuth credentials file named credentials.json in root folder
#


from __future__ import print_function
from tkinter import filedialog
import tkinter as tk
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError

import sys
import io
import os
import os.path
import pprint
import time
import math
import mimetypes

MAX_COURSENUM = 100

root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)
MODOUTFOLDER = filedialog.askdirectory() + '/courses/'
if(MODOUTFOLDER == '/courses/'):
	exit(1)
print("Saving to: "+MODOUTFOLDER)

INVALID_FILE = '<>:"/\|?* '

def validify(filename):
	for char in INVALID_FILE:
		filename = filename.replace(char, '_')
	return filename

def printenc(*value, sep=' ', end='\n', file=sys.stdout, flush=False):
	print(*[x.encode(file.encoding, errors='replace').decode(file.encoding,
	      errors='replace') for x in value], sep=sep, end=end, file=file, flush=flush)

def pprintenc(object, stream = sys.stdout, indent = 1, width = 80, depth = None, *, compact = False, sort_dicts = True, underscore_numbers = False):
	pprint.pprint([x.encode(stream.encoding, errors='replace').decode(stream.encoding,
	      errors='replace') for x in object], stream=stream, indent=indent, width=width, depth=depth, compact=compact, sort_dicts=sort_dicts, underscore_numbers=underscore_numbers)

def readBlacklist():

	with open(MODOUTFOLDER + "course_blacklist.txt", "a+", encoding='utf-8') as fd:
		fd.seek(0)
		course_blacklist = fd.read().splitlines()
	return course_blacklist

def main():
	os.makedirs(os.path.dirname(MODOUTFOLDER), exist_ok=True)

	service = retrieve_service()
	courses_ = service.courses().list(pageSize=MAX_COURSENUM).execute()
	

	courses = courses_.copy()		
	course_blacklist = readBlacklist()
	courses['courses'] = [x for x in courses_['courses'] if x['name'] not in course_blacklist]

	printenc("\n"+"#"*11)
	printenc(f"Course list")
	printenc("#"*11+"\n")
	for course in courses['courses']:
		printenc(course['name'])

	input("\nUpdate courses blacklist and press Enter to continue...")

	courses = courses_.copy()		
	course_blacklist = readBlacklist()
	courses['courses'] = [x for x in courses_['courses'] if x['name'] not in course_blacklist]

	printenc("\n"+"#"*19)
	printenc(f"Updated Course list")
	printenc("#"*19+"\n")
	for course in courses['courses']:
		printenc(course['name'])


	downloaded_files = list()
	skipped_files = list()
	failed_files = list()

	for course in courses['courses']:

		course_name = validify(course['name'])
		course_id = course['id']
		printenc("\n"+"#"*(len(course_name)+22))
		printenc(f"Downloading files for {course_name}")
		printenc("#"*(len(course_name)+22)+"\n")
		if not (os.path.exists(course_name)):
			os.makedirs(MODOUTFOLDER + course_name, exist_ok=True)

		announcements = service.courses().announcements().list(courseId=course_id).execute()
		workmaterials = service.courses().courseWorkMaterials().list(courseId=course_id).execute()
		work = service.courses().courseWork().list(courseId=course_id).execute()

		download, skipped, failed = download_announcement_files(announcements, course_name)
		downloaded_files = downloaded_files + download
		skipped_files = skipped_files + skipped
		failed_files = failed_files + failed
		
		download, skipped, failed = download_workmater_files(workmaterials, course_name)
		downloaded_files = downloaded_files + download
		skipped_files = skipped_files + skipped
		failed_files = failed_files + failed

		download, skipped, failed = download_works_files(work, course_name)
		downloaded_files = downloaded_files + download
		skipped_files = skipped_files + skipped
		failed_files = failed_files + failed

	with open(MODOUTFOLDER + 'DOWNLOADED.txt', 'w', encoding='utf-8') as downloaded:
		pprintenc(downloaded_files, downloaded, width=200)
	with open(MODOUTFOLDER + 'SKIPPED_DOWNLOADS.txt', 'w', encoding='utf-8') as sdownloaded:
		pprintenc(skipped_files, sdownloaded, width=200)
	with open(MODOUTFOLDER + 'FAILED_DOWNLOADS.txt', 'w', encoding='utf-8') as fdownloaded:
		pprintenc(failed_files, fdownloaded, width=200)


def retrieve_service():

	SCOPES = [
		'https://www.googleapis.com/auth/classroom.courses.readonly',
		'https://www.googleapis.com/auth/classroom.announcements.readonly',
		'https://www.googleapis.com/auth/classroom.student-submissions.me.readonly',
		'https://www.googleapis.com/auth/classroom.courseworkmaterials.readonly'
	]
	credentials = None
	# token-classroom.json - classrooms tokens

	if os.path.exists(MODOUTFOLDER + 'token-classroom.json'):
		credentials = Credentials.from_authorized_user_file(
			MODOUTFOLDER + 'token-classroom.json', SCOPES)

	# If there are no (valid) credentials available, let the user log in.
	if not credentials or not credentials.valid:
		if credentials and credentials.expired and credentials.refresh_token:
			credentials.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				'credentials.json', SCOPES)
			credentials = flow.run_local_server(port=0)
		# Save the credentials for the next run
		with open(MODOUTFOLDER + 'token-classroom.json', 'w') as token:
			token.write(credentials.to_json())

	try:
		return build('classroom', 'v1', credentials=credentials)

	except HttpError as error:
		printenc('>> An error occurred on service creation: %s' % error)


google_mimetypes = {'application/vnd.google-apps.document': ['application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.docx'], 'application/vnd.google-apps.drawing': ['application/pdf', '.pdf'], 'application/vnd.google-apps.presentation': ['application/vnd.openxmlformats-officedocument.presentationml.presentation', '.pptx'], 'application/vnd.google-apps.spreadsheet': ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.xlsx']}
def resolveGoogleMime(mimetype):
	if mimetype in google_mimetypes:
		return google_mimetypes[mimetype]
	return None

def resolveFileName(file_id):
	SCOPES = ['https://www.googleapis.com/auth/drive']

	credentials = None
	# token-drive.json - drive tokens

	if os.path.exists(MODOUTFOLDER + 'token-drive.json'):
		credentials = Credentials.from_authorized_user_file(
			MODOUTFOLDER + 'token-drive.json', SCOPES)

	# relogin
	if not credentials or not credentials.valid:
		if credentials and credentials.expired and credentials.refresh_token:
			credentials.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
			credentials = flow.run_local_server(port=0)
		# Save credentials
		with open(MODOUTFOLDER + 'token-drive.json', 'w') as token:
			token.write(credentials.to_json())

	service = build('drive', 'v3', credentials=credentials)
	
	
	try:
		file_ = service.files().get(fileId=file_id).execute()
	except HttpError as error:
		printenc(f'>> An error occurred while resolving the filename: {error}')
		return None
	filename = validify(file_['name'])


	if file_['mimeType'] == "":
		return filename

	extr = resolveGoogleMime(file_['mimeType'])
	if extr is None:
		ext = mimetypes.guess_extension(file_['mimeType'])
		guess = mimetypes.guess_type('a'+os.path.splitext(filename)[-1])[0]
		if ext is not None and (guess is None or guess != file_['mimeType']):
			return filename + ext
	else:
		return ''.join(os.path.splitext(filename)[:-1]) + extr[1]

	return filename

def download_file(file_id, file_name, course_name):

	SCOPES = ['https://www.googleapis.com/auth/drive']
	
	failed_downloads = None

	credentials = None

	# token-drive.json - drive tokens
	if os.path.exists(MODOUTFOLDER + 'token-drive.json'):
		credentials = Credentials.from_authorized_user_file(MODOUTFOLDER + 'token-drive.json', SCOPES)
	
	# relogin
	if not credentials or not credentials.valid:
		if credentials and credentials.expired and credentials.refresh_token:
			credentials.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
			credentials = flow.run_local_server(port=0)
		# Save credentials
		with open(MODOUTFOLDER + 'token-drive.json', 'w') as token:
			token.write(credentials.to_json())
	
	service = build('drive', 'v3', credentials=credentials)
	try:
		file_ = service.files().get(fileId=file_id, fields="*").execute()
		mimetype = file_['mimeType']
		mtype = resolveGoogleMime(mimetype)
		# try export_media for google mime-types
		if mtype is not None:
			mimetype = mtype[0]
			request = service.files().export_media(fileId=file_id, mimeType=mtype[0])
		# try get_media for known mime-types
		else:
			request = service.files().get_media(fileId=file_id)

		fd = io.BytesIO()
		downloader = MediaIoBaseDownload(fd, request)
		
		dl = False
		while not dl:
			status, dl = downloader.next_chunk()
			printenc(f'Download {int(status.progress() * 100)}% - [{mimetype}]')

		fd.seek(0)

		with open(os.path.join(MODOUTFOLDER, course_name, file_name), 'wb') as fo:
			fo.write(fd.read())
			fo.close()
	except HttpError as error:
		# retry with lfs
		if error.error_details[0]['reason'] == "exportSizeLimitExceeded" and mtype is not None:
			printenc(f'>> Download failed, please check FAILED_DOWNLOADS.txt')
			file_exportlinks = file_['exportLinks'].values()
			failed_downloads = file_name + ': [' + ', '.join(file_exportlinks) + ']'

		else:
			failed_downloads = file_name + ': [unknown]'
			printenc(f'>> An error occurred while downloading: {error}')

	return failed_downloads

def extract_id(link):
	return link[link.rfind('/d/')+3:link.rfind('/')]
def extract_exportId(link):
	return link[link.rfind('?id=')+4:link.rfind('&')]

def resolve_idAndName(material):
	file_id = material['driveFile']['driveFile']['id']
	file_name = material['driveFile']['driveFile']['title']

	# using alterative link on templates
	if (file_name[0:10] == "[Template]"):
			file_altern_link = material['driveFile']['driveFile']['alternateLink']
			file_id = extract_id(file_altern_link)
			printenc(f"<< Downloading {file_id} from {file_altern_link} instead")

	file_name = resolveFileName(file_id)
	if file_name is None:
		file_name = material['driveFile']['driveFile']['title']

	return file_id, file_name

def download_announcement_files(announcements, course_name):
	announcement_list = announcements.keys()
	downloaded = list()
	failed_downloads = list()
	skipped_downloads = list()
	if (len(announcement_list) != 0):
		existing_files = getDirents(os.path.join(MODOUTFOLDER, course_name))

		for announcement in announcements['announcements']:
			try:  # if this announcements contain a file then do this
				for material in announcement['materials']:
					
					file_id, file_name = resolve_idAndName(material)

					path_str = os.path.join(MODOUTFOLDER, course_name, file_name)
					
					if file_name not in existing_files:
						printenc("DOWNLOADING Announcement:", file_name)
						dllfail = download_file(file_id, file_name, course_name)
						if dllfail is None:
							downloaded.append("Announcement:  "+course_name + ': ' + file_name)
						else:
							failed_downloads.append("Announcement:  "+course_name + ': ' + dllfail)
					elif file_name in existing_files:
						skipped_downloads.append("Announcement: "+course_name + ': ' + file_name)
						printenc(file_name, "Already exists - skipping")
					else:
						failed_downloads.append("Announcement:  "+course_name + ': ' + file_name)
						printenc("Could not create file")
			except KeyError as e:
				printenc(f">> KeyError {e}")
	return downloaded, skipped_downloads, failed_downloads


def download_workmater_files(workmaterials, course_name):
	workmater_list = workmaterials.keys()
	downloaded = list()
	failed_downloads = list()
	skipped_downloads = list()
	if (len(workmater_list) != 0):
		existing_files = getDirents(os.path.join(MODOUTFOLDER, course_name))

		for material in workmaterials['courseWorkMaterial']:
			try:  # if this material contains a file then do this
				for material in material['materials']:
					
					file_id, file_name = resolve_idAndName(material)

					path_str = os.path.join(MODOUTFOLDER, course_name, file_name)

					if file_name not in existing_files:
						printenc("DOWNLOADING Material:", file_name)
						dllfail = download_file(file_id, file_name, course_name)
						if dllfail is None:
							downloaded.append("WorkMaterial:  "+course_name + ': ' + file_name)
						else:
							failed_downloads.append("WorkMaterial:  "+course_name + ': ' + dllfail)
					elif file_name in existing_files:
						skipped_downloads.append("WorkMaterial: "+course_name + ': ' + file_name)
						printenc(file_name, "Already exists - skipping")
					else:
						failed_downloads.append("WorkMaterial:  "+course_name + ': ' + file_name)
						printenc("Could not create file")
			except KeyError as e:
				continue
	return downloaded, skipped_downloads, failed_downloads


def download_works_files(works, course_name):
	works_list = works.keys()
	downloaded = list()
	failed_downloads = list()
	skipped_downloads = list()
	if (len(works_list) != 0):
		existing_files = getDirents(os.path.join(MODOUTFOLDER, course_name))

		for work in works['courseWork']:
			try:  # if this announcements contain a file then do this
				for material in work['materials']:
					
					file_id, file_name = resolve_idAndName(material)

					path_str = os.path.join(MODOUTFOLDER, course_name, file_name)
					
					if file_name not in existing_files:
						printenc("DOWNLOADING Work:", file_name)
						dllfail = download_file(file_id, file_name, course_name)
						if dllfail is None:
							downloaded.append("Work:  "+course_name + ': ' + file_name)
						else:
							failed_downloads.append("Work:  "+course_name + ': ' + dllfail)
					elif file_name in existing_files:
						skipped_downloads.append("WorkMaterial: "+course_name + ': ' + file_name)
						printenc(file_name, "Already exists - skipping")
					else:
						failed_downloads.append("Work:  "+course_name + ': ' + file_name)
						printenc("Could not create file")

			except KeyError as e:
				continue
	return downloaded, skipped_downloads, failed_downloads

def getDirents(dirName):
	dirents = os.listdir(dirName)
	files = list()
	# all entries
	for filent in dirents:
		# Create path
		path = os.path.join(dirName, filent)
		# recursively search subdir
		if os.path.isdir(path):
			files = files + getDirents(path)
		else:
			files.append(path)
	# return only the names
	return [fi[fi.rfind('\\')+1:] for fi in files]


if __name__ == '__main__':
	start = time.time()
	main()
	tk.messagebox.showinfo("ClassroomDownloader by Felix Kröhnert", "Completed successfully in "+str(math.modf(time.time()-start)[1])+"s\nSaved to: "+MODOUTFOLDER)
	print(f'\nRuntime: {time.time()-start}')
