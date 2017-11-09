import lxml.html as lh
from selenium import webdriver
import time,openpyxl,re,os
#fileName1="Ubuntu_100_Link_FreeLancer_sample01.xlsx"
#pathToFile= os.path.dirname(os.path.realpath(__file__))+str("\\") + fileName
pathToFile= r"F:\common\scraping\UbuntuForumScrape\com\ubuntu\Ubuntu_100_Link_FreeLancer_sample01.xlsx"
excelPath=pathToFile
SheetName='100 Threads  2-15 Thabit'
workbook = openpyxl.load_workbook(excelPath)
worksheet = workbook.get_sheet_by_name(SheetName)
website = []
for row in worksheet.iter_rows(min_row=2,max_row=103, min_col=3, max_col=3):  ##Getting all the websites,102 max row
    for cell in row:
        website.append(cell.value)

#browser=webdriver.Chrome(r'C:\Users\abhis\Pycharmojects\webdriver\chromedriver_win32\chromedriver')
USEREMAIL=''    #provide email which was used for registration
userpassword='' #provide password which was used for registration
browser=webdriver.PhantomJS(r'C:\Users\abhis\PycharmProjects\webdriver\phantomjs-2.1.1-windows\phantomjs-2.1.1-windows\bin\phantomjs')
class UbuntuForum(object):
    def __init__(self):
        self.ubuntuUrl='https://ubuntuforums.org/'
        browser.get("https://login.ubuntu.com/")
        time.sleep(2)  # Let the user actually see something!
        search_box_username = browser.find_element_by_xpath('//*[@id="id_email"]')
        time.sleep(2)
        search_box_username.send_keys(USEREMAIL)
        search_box_password = browser.find_element_by_xpath('//*[@id="id_password"]')
        search_box_password.send_keys(userpassword)
        search_box_password.submit()
        time.sleep(3)  # Let the user actually see something!
        browser.get("https://ubuntuforums.org/showthread.php?t=756434")
        # time.sleep(2)
        browser.find_element_by_xpath('//input[@value="Login with SSO"]').click()
        time.sleep(2)
        # LogMeIn
        browser.find_element_by_xpath('//*[@id="auth"]/div[2]/form/p/button/span/span').submit()
        time.sleep(3)
        browser.get("https://ubuntuforums.org/showthread.php?t=756434")
        time.sleep(2)
        #Variables
        self.Title = {}
        self.UserName = {}
        self.Dates = {}
        self.Location = {}
        self.Beans = {}
        self.InitialPost={}
        #self.InitialRep1 = {}
        #OtheUsers
        # self.User1={}
        # self.Date1={}
        # self.Location1={}
        # self.Beans1={}
        self.UserQuotation1={}
        self.BlockQuotation1={}
        self.Code1={}
        self.InLink1={}
        self.firstPageCol=10
        self.nextPageCol=91
        self.firstPageUserCountStartIndex=1
        self.nextpageUserCountStartIndex=11
        self.startIndex = 2
        self.nextUserIndex=11
        #self.Rep1={}#InitialPost retrieves it
        #Xpath Variables
        self.xPathVar1='//a[@class="postcounter" and contains(text(),"#'
        self.forFirstUser = {}
        for webcount in range(0,len(website)):#starts with 0 and 2 has pagination
            time.sleep(2)
            browser.get(website[webcount])#gets it from excel file , the website
            content = browser.page_source
            tree = lh.fromstring(content)
            print("I am here\n", website[webcount])
            self.dataExtraction(webcount,tree,self.firstPageUserCountStartIndex)#self.firstPageUserCountStartIndex=1
            print("First user data inserted 1")
            self.writeToExcelForFirstUser(webcount)
            print("Other user data inserted 1")
            self.writeToExcelForOtherUser(webcount,tree,self.startIndex,self.firstPageCol)#self.startIndex = 2 self.firstPageCol=10
            if not tree.xpath('//span[@class="prev_next"]/a[contains(@title,"Next Page")]'):
                pass
            else:
                NextPage = tree.xpath('//span[@class="prev_next"]/a[contains(@title,"Next Page")]/@href')
                print("Nextpage href",NextPage)
                browser.get(self.ubuntuUrl + NextPage[0])  # NextPage
                content = browser.page_source
                tree1 = lh.fromstring(content)
                # import pdb
                # pdb.set_trace()
                self.dataExtraction(webcount, tree1,self.nextpageUserCountStartIndex)#self.nextpageUserCountStartIndex=11 self.firstPageUserCountStartIndex=1
                #print("First user data inserted 1")
                #self.writeToExcelForFirstUser(webcount)
                print("Other user data inserted 1")
                self.writeToExcelForOtherUser(webcount, tree1,self.nextUserIndex,self.nextPageCol)#self.nextUserIndex=11 self.nextPageCol=82

    def dataExtraction(self,webcount,tree,startUserIndex):
        self.Title[1] = tree.xpath('//span[@class="threadtitle"]/a/text()')
        CountPage1 = tree.xpath('//span[@class="date"]/text()')

        for elem in range(startUserIndex, startUserIndex+len(CountPage1)):  # post #1 to #10
            print("ElementSequence is ", elem)
            UserName= tree.xpath(self.xPathVar1 + str(elem) + '")]/../../..//strong/span/text()')
            self.UserName[elem]=UserName[0]
            Dates= tree.xpath(self.xPathVar1 + str(elem) + '")]/../../..//span[@class="date"]/text()')
            self.Dates[elem]=Dates[0]
            ##Location
            ##//a[@class="postcounter" and contains(text(),"1")]/../../../..//dt['2' and contains(text(),"Location")]
            if not tree.xpath(self.xPathVar1 + str(elem) + '")]/../../../..//dt[2 and contains(text(),"Location")]/text()'):
                print("Location not present")
                print("no")
                self.Location[elem] = 'no'
            else:
                self.Location[elem] = tree.xpath(
                    self.xPathVar1 + str(elem) + '")]/../../../..//dd[2]/text()')
            # Beans
            ##// a[ @class ="postcounter" and contains(text(), "1")] /../../../..// dt[3 and contains(text(), "Beans")]
            ##// a[ @class ="postcounter" and contains(text(), "1")] /../../../..// dd[3]
            ##// a[ @class ="postcounter" and contains(text(), "1")] /../../../..//following-sibling::dt[contains(text(),'Beans')]//following-sibling::dd
            ##Beans=[]
            if not tree.xpath(self.xPathVar1 + str(elem) + '")]/../../../..//dt[3 and contains(text(),"Beans")]/text()'):
                print("No Beans")
                self.Beans[elem] = 'No'
            else:
                Beans = tree.xpath(self.xPathVar1 + str(elem) + '")]/../../../..//following-sibling::dt[contains(text(),"Beans")]//following-sibling::dd/text()')
                self.Beans[elem] = Beans[0]
            ##//a[@class="postcounter" and contains(text(),"4")]/../../../..//following-sibling::blockquote[@class="postcontent restore"]
            ##//a[@class="postcounter" and contains(text(),"4")]/../../../..//blockquote[@class="postcontent restore" ]//following-sibling::*[ not(@class="message")]
            ##self.xPathVar1 + str(elem) + '")]/../../../..//blockquote[@class="postcontent restore"]')
            Rep1 = tree.xpath(
                self.xPathVar1 + str(elem) + '")]/../../../..//blockquote[@class="postcontent restore"]')
            ##// blockquote[ @class ="postcontent restore"]
            ##for i in range(len(UserName)):
            ##InitialPost1=re.sub(r'[\n\t]','',(InitialPost.text_content()))
            self.InitialPost[elem] = re.sub('[\n\t]', '', Rep1[0].text_content())

            ##self.InitialRep1[elem]=(re.sub(r'[\n\t]','',Rep1))

            ##User_QuotationUser=tree.xpath('//a[@class="postcounter" and contains(text(),"#6")]/../../../..//div[@class="message"]/..//strong/text()')
            ## User_QuotationMessage=tree.xpath('//a[@class="postcounter" and contains(text(),"#6")]/../../../..//div[@class="message"]/text()')
            User_QuotationUser = []
            User_QuotationMessage = []
            if not tree.xpath(self.xPathVar1 + str(elem) + '")]/../../../..//div[@class="message"]/text()'):
                print("No")
                self.UserQuotation1[elem] = 'No'
            else:
                User_QuotationUser = tree.xpath(
                    self.xPathVar1 + str(elem) + '")]/../../../..//div[@class="message"]/..//strong/text()')
                User_QuotationMessage = tree.xpath(
                    self.xPathVar1 + str(elem) + '")]/../../../..//div[@class="message"]/text()')
                self.UserQuotation1[elem] = User_QuotationUser[0] + ':' + User_QuotationMessage[0]
            BlockQuotation2 = []
            BlockQuotation2 = tree.xpath(
                self.xPathVar1 + str(
                    elem) + '")]/../../../..//blockquote/div[2]//div[@class="quote_container"]//text()')  # need to use text function
            ##if not self.xPathVar1+str(elem)+'")]/../../../..//blockquote/div[2]//div[@class="quote_container"]/text()':
            if len(BlockQuotation2) == 0:
                self.BlockQuotation1[elem] = "No"
            else:
                self.BlockQuotation1[elem] = BlockQuotation2

            # self.Code1 = {}
            Code = []
            Code = tree.xpath(
                self.xPathVar1 + str(elem) + '")]/../../../..//*[@class="bbcode_description"]/text()')
            if tree.xpath(self.xPathVar1 + str(elem) + '")]/../../../..//*[@class="bbcode_description"]/text()'):
                if Code[0] == "Code:":
                    self.Code1[elem] = 'Yes'
                else:
                    self.Code1[elem] = 'No'
            else:
                self.Code1[elem] = 'No'
            ##Inlink=[]
            ##self.InLink1 = {}  # If there is a link, then yes else no
            ##InLink=tree.xpath(self.xPathVar1+str(elem)+'")]/../../../..//div[@class="content"]//*[@class="postcontent restore"]//a[@target="_blank"]/text()')
            if not tree.xpath(self.xPathVar1 + str(
                    elem) + '")]/../../../..//div[@class="content"]//*[@class="postcontent restore"]//a[@target="_blank"]/text()'):
                self.InLink1[elem] = 'No'
            else:
                self.InLink1[elem] = 'Yes'
            # self.InLink1=//a[@class="postcounter" and contains(text(),"#6")]/../../../..//blockquote/div[2]//div[@class="bbcode_quote_container"]//following-sibling::a

            self.forFirstUser = {
                1: re.sub(r"[\[\"*\"\]]", '', str(self.Title[1]))
                , 2: re.sub(r"[\[\'*\'\]]", '', str(self.UserName[1]))
                , 3: re.sub(r"[\[\'*\'\]]", '', str(self.Dates[1]))
                , 4: re.sub(r"[\[\'*\'\]]", '', str(self.Location[1]))
                , 5: self.Beans[1]
                , 6: self.InitialPost[1]
            }
        print("Title", self.Title
              , "UserName", self.UserName
              , "Dates1", self.Dates
              , "Location1", self.Location
              , "Beans", self.Beans
              , "InitialPost", self.InitialPost
              , "self.UserQuotation1", self.UserQuotation1
              , "self.BlockQuotation1", self.BlockQuotation1
              , "self.Code1", self.Code1
              , "self.InLink1", self.InLink1
              , end='\n', sep='|')

    def writeToExcelForFirstUser(self, webcount):
        for index in range(1, 7):
            col = 3
            row = webcount + 2
            try:
                worksheet.cell(row=row, column=col + index).value = str(self.forFirstUser[index])
            except KeyError as e:
                print("Unable to print value")
        print("First user data inserted")

    def writeToExcelForOtherUser(self, webcount, tree,startUserIndex,startpageCol):
        #startIndex = 2
        CountPage2 = tree.xpath('//span[@class="date"]/text()')
        for index, col in zip(range(startUserIndex, startUserIndex+len(CountPage2) ),
                              range(startpageCol, startpageCol +len(CountPage2) * 9, 9)):  # index needs to start with 2 for 2nd user onwards
            # col=10
            print("writeToExcelForOtherUser",index,col)
            row = webcount + 2
            try:
                worksheet.cell(row=row, column=col + 0).value = re.sub(r"[\[\'*\'\]]", '', str(self.UserName[index]))
                worksheet.cell(row=row, column=col + 1).value = re.sub(r"[\[\'*\'\]]", '', str(self.Dates[index]))
                worksheet.cell(row=row, column=col + 2).value = re.sub(r"[\[\'*\'\]]", '', str(self.Location[index]))
                worksheet.cell(row=row, column=col + 3).value = str(self.Beans[index])
                worksheet.cell(row=row, column=col + 4).value = str(self.UserQuotation1[index])
                worksheet.cell(row=row, column=col + 5).value = re.sub(r'[\n\t]*', '',
                                                                       ''.join(self.BlockQuotation1[index]))
                worksheet.cell(row=row, column=col + 6).value = str(self.Code1[index])
                worksheet.cell(row=row, column=col + 7).value = str(self.InLink1[index])
                worksheet.cell(row=row, column=col + 8).value = str(self.InitialPost[index])
            except KeyError as e:
                print("Unable to print value becuase of %s and at index %d" % (e, index))
        print("First user data inserted")

if __name__ == '__main__':
    print("Start")
    obj = UbuntuForum()
    # obj.main()
    #obj.dataExtraction()
    # processingFirstPost()
    print("The End")
    workbook.save(excelPath)
    browser.quit()
