# author: Nolan B. Kennedy (nxkennedy)

import csv
import os
import time
import sqlite3

rootdir = os.getcwd()
components = []
reports = []
dbname = 'db-' + time.strftime("%Y%m%d-%H%M%S") + '.db'


def normalize(infile):

	filename = os.path.basename(infile)
	info = filename.split(' ')
	s_id = info[2]
	s_type = info[6]
	if s_type == 'Data':
		s_type = 'Data Entry'
	normalized = []

	# This is how we normalize our reports
	master = {
		# new fields that didn't exist
		'scenario_id': s_id,
		'scenario_type': s_type,
		'component': 'NA',
		# shared by all
		'Email': 'NA',
		'Recipient Name': 'NA',
        	'Recipient Group': 'NA',
		'Department': 'NA',
		'Location': 'NA',
		'Opened Email?': 'NA',
		'Opened Email Timestamp': 'NA',
		# specific to attachment
		'Viewed Education?': 'NA',
		'Viewed Education Timestamp': 'NA',
		# shared by clicked link, data entry
		'Clicked Link?': 'NA',
		'Clicked Link Timestamp': 'NA',
		# specific to data entry
		'Submitted Form': 'NA',
		'Username': 'NA',
		'Entered Password?': 'NA',
		'Submitted Form Timestamp': 'NA',
		# shared by all
		'Reported Phish?': 'NA',
		'New/Repeat Reporter': 'NA',
		'Reported Phish Timestamp': 'NA',
		'Time to Report (in seconds)': 'NA',
		'Remote IP': 'NA',
		'GeoIP Country': 'NA',
		'GeoIP City': 'NA',
		'GeoIP Organization': 'NA',
		'Last DSN': 'NA',
		'Last Email Status': 'NA',
		'Last Email Status Timestamp': 'NA',
		'Language': 'NA',
		'Browser': 'NA',
		'User-Agent': 'NA',
		'Mobile?': 'NA',
		# specific to clicked link
		'Seconds Spent on Education Page': 'NA',
		'ACCOUNTTYPE': 'NA', # a key that you will likely have to change here and twice in the create_db() function
		'SN': 'NA',
		'GIVENNAME': 'NA',
		'INITIALS': 'NA',
		# specific to data entry
		'DUTY_STATION': 'NA',
		'RAND': 'NA',
		'Submitted Data': 'NA'}

	with open(infile) as csvfile:
		reader = csv.DictReader(csvfile)
		for row in reader:
			for key, value in row.items():

				if key in master:

					if value == '':
						master[key] = 'NULL'
					else:
						master[key] = value
			for c in components:
				if c in master['Recipient Group']:
					master['component'] = c

			normalized.append(master.copy())
		return normalized


def create_db():

	print('[+] Normalizing and inserting reports into a database...')
	fieldnames = [
		# new fields that didn't exist
		'scenario_id', 'scenario_type', 'component',
		# shared by all
		'Email', 'Recipient Name', 'Recipient Group', 'Department', 'Location', 'Opened Email?', 'Opened Email Timestamp',
		# specific to attachement
		'Viewed Education?', 'Viewed Education Timestamp',
		# shared by clicked link, data entry
		'Clicked Link?', 'Clicked Link Timestamp',
		# specific to data entry
		'Submitted Form', 'Username', 'Entered Password?', 'Submitted Form Timestamp',
		# shared by all
		'Reported Phish?', 'New/Repeat Reporter', 'Reported Phish Timestamp', 'Time to Report (in seconds)', 'Remote IP',
		'GeoIP Country', 'GeoIP City', 'GeoIP Organization', 'Last DSN', 'Last Email Status',
		'Last Email Status Timestamp', 'Language', 'Browser', 'User-Agent', 'Mobile?',
		# specific to clicked link
		'Seconds Spent on Education Page', 'ACCOUNTTYPE', 'SN', 'GIVENNAME', 'INITIALS',
		# specific to data entry
		'DUTY_STATION', 'RAND', 'Submitted Data']

	db = sqlite3.connect(dbname)
	c = db.cursor()
	c.execute('''CREATE TABLE Reports
			('scenario_id' INTEGER,
			'scenario_type' TEXT,
			'component' TEXT,
			'Email' TEXT,
			'Recipient Name' TEXT,
        		'Recipient Group' TEXT,
			'Department' TEXT,
			'Location' TEXT,
			'Opened Email?' TEXT,
			'Opened Email Timestamp' TEXT,
			'Viewed Education?' TEXT,
			'Viewed Education Timestamp' TEXT,
			'Clicked Link?' TEXT,
			'Clicked Link Timestamp' TEXT,
			'Submitted Form' TEXT,
			'Username' TEXT,
			'Entered Password?' TEXT,
			'Submitted Form Timestamp' TEXT,
			'Reported Phish?' TEXT,
			'New/Repeat Reporter' TEXT,
			'Reported Phish Timestamp' TEXT,
			'Time to Report (in seconds)' TEXT,
			'Remote IP' TEXT,
			'GeoIP Country' TEXT,
			'GeoIP City' TEXT,
			'GeoIP Organization' TEXT,
			'Last DSN' TEXT,
			'Last Email Status' TEXT,
			'Last Email Status Timestamp' TEXT,
			'Language' TEXT,
			'Browser' TEXT,
			'User-Agent' TEXT,
			'Mobile?' TEXT,
			'Seconds Spent on Education Page' TEXT,
			'ACCOUNTTYPE' TEXT,
			'SN' TEXT,
			'GIVENNAME' TEXT,
			'INITIALS' TEXT,
			'DUTY_STATION' TEXT,
			'RAND' TEXT,
			'Submitted Data' TEXT)''')

	for r in reports:
		normalized = normalize(r)
		for row in normalized:
			lst = [v for v in row.values()]
			placeholders = ",".join('?' * len(lst))
			query = "INSERT INTO Reports VALUES (%s)" % (placeholders)
			c.execute(query, lst)
			db.commit()
	db.close()

	print('[+] DONE. Database created')


def scan_directory():

	print('[+] Scanning current directory "{0}"...'.format(rootdir))
	for root, dirs, files in os.walk(rootdir):
		for dir in dirs:
			component = os.path.basename(dir)
			components.append(component)
			fullpath = os.path.join(root, dir)

		for f in files:
			fullpath = os.path.join(root, f)
			if f.endswith('.csv'):
			    reports.append(fullpath)
			else:
				print('[!] INVALID REPORT: "{0}". Skipping...'.format(fullpath))
	print('[+] Done scanning directory')




if __name__ == '__main__':
	scan_directory()
	create_db()
