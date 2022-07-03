import os
import pickle
# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode
# for dealing with attachement MIME types
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from mimetypes import guess_type as guess_mime_type
import dateutil.parser as parser
import time


# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/']
#insert your email here
our_email = 'useremail@gmail.com'

def gmail_authenticate():
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials availablle, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

# get the Gmail API service
service = gmail_authenticate()

# Adds the attachment with the given filename to the given message
def add_attachment(message, filename):
    content_type, encoding = guess_mime_type(filename)
    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(filename, 'rb')
        msg = MIMEText(fp.read().decode(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(filename, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(filename, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(filename, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    filename = os.path.basename(filename)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)
    
    
def build_message(destination, obj, body, attachments=[]):
	if not attachments: # no attachments given
		message = MIMEText(body)
		message['to'] = destination
		message['from'] = our_email
		message['subject'] = obj
	else:
		message = MIMEMultipart()
		message['to'] = destination
		message['from'] = our_email
		message['subject'] = obj
		message.attach(MIMEText(body))
	for filename in attachments:
		add_attachment(message, filename)
	return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}
    
    
    
def send_message(service, destination, obj, body, attachments=[]):
	return service.users().messages().send(
	      userId="me",
	      body=build_message(destination, obj, body, attachments)
	    ).execute()

  #sending message  
  #send_message(service, "recieveremail@gmail.com", "Hello",

 # "This is the body of the email",["attachment.file"])
while True:
	user_id =  'me'
	label_id_one = 'INBOX'
	label_id_two = 'UNREAD'

	# Getting all the unread messages from Inbox
	# labelIds can be changed accordingly
	unread_msgs = service.users().messages().list(userId='me',labelIds=[label_id_one, label_id_two]).execute()
	print(unread_msgs)
	#print("lol")
	# We get a dictonary. Now reading values for the key 'messages'
	if(unread_msgs['resultSizeEstimate'])!=0:
		mssg_list = unread_msgs['messages']
		#print(mssg_list)
		print ("Total unread messages in inbox: ", str(len(mssg_list)))

		final_list = [ ]
		sends_name = []


		for mssg in mssg_list:
			temp_dict = { }
			m_id = mssg['id'] # get id of individual message
			message = service.users().messages().get(userId=user_id, id=m_id).execute() # fetch the message using API
			payld = message['payload'] # get payload of the message 
			headr = payld['headers'] # get header of the payload


			for one in headr: # getting the Subject
				if one['name'] == 'Subject':
					msg_subject = one['value']
					temp_dict['Subject'] = msg_subject
				else:
					pass


			for two in headr: # getting the date
				if two['name'] == 'Date':
					msg_date = two['value']
					date_parse = (parser.parse(msg_date))
					m_date = (date_parse.date())
					temp_dict['Date'] = str(m_date)
				else:
					pass

			for three in headr: # getting the Sender
				if three['name'] == 'From':
					msg_from = three['value']
					f=1
					for x in sends_name:
						if x==msg_from:
							f=0
					if f==1:
						sends_name.append(msg_from)		
					temp_dict['Sender'] = msg_from
				else:
					pass

			temp_dict['Snippet'] = message['snippet'] # fetching message snippet


			try:
				
				# Fetching message body
				mssg_parts = payld['parts'] # fetching the message parts
				part_one  = mssg_parts[0] # fetching first element of the part 
				part_body = part_one['body'] # fetching body of the message
				part_data = part_body['data'] # fetching data from the body
				clean_one = part_data.replace("-","+") # decoding from Base64 to UTF-8
				clean_one = clean_one.replace("_","/") # decoding from Base64 to UTF-8
				clean_two = base64.b64decode (bytes(clean_one, 'UTF-8')) # decoding from Base64 to UTF-8
				soup = BeautifulSoup(clean_two , "lxml" )
				mssg_body = soup.body()
				# mssg_body is a readible form of message body
				# depending on the end user's requirements, it can be further cleaned 
				# using regex, beautiful soup, or any other method
				temp_dict['Message_body'] = mssg_body

			except :
				pass

			print (temp_dict)
			final_list.append(temp_dict) # This will create a dictonary item in the final list
			
			# This will mark the messagea as read
			service.users().messages().modify(userId=user_id, id=m_id,body={ 'removeLabelIds': ['UNREAD']}).execute() 
			




        #text to speech







		from gtts import gTTS
		  
		# This module is imported so that we can 
		# play the converted audio
		import os
		from playsound import playsound 
		# The text that you want to convert to audio
		#mytext = 'heyy baby!'
		senders = ""
		for r in sends_name:
			senders += r
		if len(final_list) == 1:
			mytext=" You have {} unread message from {}".format(len(final_list),senders)
		else:	
			mytext = " You have {} unread messages from {}".format(len(final_list),senders)
		language = 'en'
		# Language in which you want to convert

		  
		# Passing the text and language to the engine, 
		# here we have marked slow=False. Which tells 
		# the module that the converted audio should 
		# have a high speed
		myobj = gTTS(text=mytext, lang=language, slow=False)
		  
		# Saving the converted audio in a mp3 file named
		# welcome 
		myobj.save("welcome.mp3")
		  
		# Playing the converted file
		playsound("welcome.mp3")

		print ("Total messaged retrived: ", str(len(final_list)))
		print(final_list)
	#rerunning the program every 100 sec
	time.sleep(100)
