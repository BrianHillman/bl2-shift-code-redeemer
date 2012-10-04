import win32api, win32con,win32ui,os,time,twitter,win32clipboard,win32com.client
from win32api import GetSystemMetrics
from autopy import mouse,key,alert
import win32pdh, win32api
import pickle


tApi = twitter.Api(
consumer_key='JX4t5O7M3tphK3uyzjUg',
consumer_secret='ILgpx3VHoG3rNiUgmSQJdIsKA3z1YPF3nHberHEmOaA',
access_token_key='***********************************************************',
access_token_secret='#########################################################')

codeList = pickle.load(open("api","r"))
code = "CT5TB-H3ZX9-6TTBJ-JT33J-TFJRK"


def WindowExists(name):
	junk, instances = win32pdh.EnumObjectItems(None,None,'process', win32pdh.PERF_DETAIL_WIZARD)
	temp = os.popen("ps-Af").read()
	print temp
	if name in instances:
		return True
	else:
		return False

def addToUsedCodes(code):
	if code in codeList:
		print "oops, tried to re-mark a code as used"
		return
	print "'using':",code
	codeList.append(code)
	pickle.dump(codeList,open("api","w"))


def parseTweet(message):
	if message.find("PC") != -1:
		message = message[message.find("PC"):]
	elif message.find("pc") != -1:
		message = message[message.find("pc"):]
	else:
		return ""
	return message[message.find("-")-5:message.find("-")+24].strip()


def copyToClipboard(message):
	#os.system('echo ' + message.strip() + '| clip')
	win32clipboard.OpenClipboard()
	win32clipboard.EmptyClipboard()
	win32clipboard.SetClipboardText(message)
	win32clipboard.CloseClipboard()

def launchBl():
	os.system("cd C:\\Program Files (x86)\\Steam\\  && steam.exe -applaunch 49520 -NoLauncher"  )
	

def click(x,y):
	mouse.move(x, y)
	time.sleep(.5)
	mouse.click()



def clickCenter():
	click(int(.5*GetSystemMetrics (0)),int(.5*GetSystemMetrics (1)))
	time.sleep(1)

def clickExtras():
	click(int(.10468521229*GetSystemMetrics (0)),int(.4921875*GetSystemMetrics (1)))
	time.sleep(1)

def clickShift():
	click(int(.11786237188*GetSystemMetrics (0)),int(.140625*GetSystemMetrics (1)))
	time.sleep(1)
def enterCode():
	click(int(.3806734992*GetSystemMetrics (0)),int(.85286458333*GetSystemMetrics (1)))
	time.sleep(1)
	key.toggle( key.K_CONTROL, True)
	time.sleep(.1)
	key.tap( 'v' )
	key.toggle( key.K_CONTROL, False)
	time.sleep(1)
	click(int(.68008784773*GetSystemMetrics (0)),int(.571614583*GetSystemMetrics (1)))
addToUsedCodes(code)
timeline = tApi.GetUserTimeline("DuvalMagic")
code = parseTweet(timeline[0].text)

while 1:
	

	timeline = tApi.GetUserTimeline("DuvalMagic")

	tweet = timeline[0].text
	lastTweet = code
	code =  parseTweet(tweet)

	
	if not code in codeList:

		print code
		if code != "":
			copyToClipboard(code)

			if  not WindowExists("Borderlands2"):

				time.sleep(60*40)
				copyToClipboard(code)
				launchBl()
				time.sleep(60)
				#click center  for fun / make sure window's on top
				clickCenter()
				clickCenter()
				clickExtras()
				clickShift()
				enterCode()
				time.sleep(10)
				os.system("TASKKILL /IM Borderlands2.exe")
				addToUsedCodes(code)
			else:
				alert.alert(title = "New BL2 shift code! \nyou seem to be running bl2 so i copied it to your clipboard.", msg="bl2 shift code", default_button="OK")
				addToUsedCodes(code)


	time.sleep(300)




