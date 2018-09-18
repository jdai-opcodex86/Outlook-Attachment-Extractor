import win32com.client
import os,sys,argparse,time
import datetime as dt



def initialization():
	try:
		print ("Opening Outlook...")
		outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace('MAPI')
		for i in range(100):
			try:
				displayfolderselection = outlook.GetDefaultFolder(i)
				name = displayfolderselection.Name
				print(i, name)
			except:
				pass
	except:
		print ("Could not open Outlook")
		sys.exit(1)


		
def enumeratefolder(createpath,folders):
	try:
		if not os.path.exists(createpath):
			os.makedirs(createpath)
		total = folders.Count
		print("Found {} sub-folders".format(total))
		for folder in folders:
			print("Working on {} folder..".format(folder))
			foldermessages = folder.Items
			directory = createpath+"\\"+folder.Name
			print("Assigned {} folder".format(directory))
			extractattachments(directory,foldermessages)
			if total > 0:
				nextfoldername = folder.Name
				nextfolder = folder.Folders
				nextdirectory = createpath +'\\'+nextfoldername
				enumeratefolder(nextdirectory,nextfolder)
	except Exception as e:
		print(e)
		pass	
		
def extractattachments(extractpath,messages):
	try:
		if not os.path.exists(extractpath):
			os.makedirs(extractpath)
		for message in messages:
			attachments = message.Attachments
			TotalAttachments = attachments.Count
			if TotalAttachments > 0:
				print("Found %d attachments" % TotalAttachments)
				for attachment in attachments:
					date_time = time.strftime('%m-%d-%Y')
					attachment_name = date_time + '_' + attachment.FileName
					savepath = extractpath+'\\'+str(attachment_name)
					savingmessage = "Saving {} to {}"
					print(savingmessage.format(str(attachment_name), savepath))  
					attachment.SaveAsFile(savepath)
	
	except Exception as e:
		print(e)
		pass
		
def main(target,filepath):
	try:
		outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace('MAPI')
		inbox = outlook.GetDefaultFolder(target)
		messages = inbox.Items
		foldername = inbox.Name
		folders = inbox.Folders
		directory = filepath+foldername
		extractattachments(directory,messages)
		enumeratefolder(directory,folders)
	except Exception as e:
		print(e)
		pass

if __name__ == "__main__":
	initialization()
	parser = argparse.ArgumentParser()
	parser.add_argument('-t', type=int, help="Enter the number for the desired folder to extract (i.e. 6)", action="store", dest='target', required=True)
	parser.add_argument('-p', help="File Path to save the attachments (i.e. C:\\Outlook_Attachments\\", action="store", dest='filepath', required=True)
	args = parser.parse_args()
	main(args.target,args.filepath)

