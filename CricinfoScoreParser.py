'''
Parse cricket score and notify if a wicket falls or on boundary

Run the script and Enter URL from cricinfo in the text box provided and press start button.
Ex URL : 'http://www.espncricinfo.com/series/8726/game/1157486/eastern-province-vs-boland-pool-a-csa-3-day-provincial-cup-2018-19'
'''

import requests
from lxml.html import fromstring
import wx
#import winsound 
from win10toast import ToastNotifier
import time

def GetScore(link):
    http_proxy  = "http://xxx.xxx.0.25:xxxx"
    https_proxy = "https://xxx.xxx.0.25:xxxx"   
    
    proxyDict = { 
                  "http"  : http_proxy, 
                  "https" : https_proxy,               
                }    
    s = requests.Session()        
#    s.proxies = proxyDict 
    
    if link == '':
        link = 'http://www.espncricinfo.com/series/8044/game/1152562/hobart-hurricanes-vs-melbourne-renegades-52nd-match-big-bash-league-2018-19' 
        
    r = s.get(link)
    tree = fromstring(r.content)
    score = tree.findtext('.//title')
    return score


def ShowNotification(title, score, delay =2):
    toaster = ToastNotifier()
    toaster.show_toast(title,
               score,
               icon_path=None,
               duration=delay,
               threaded=True)
    while toaster.notification_active(): time.sleep(0.1)
 
class MyForm(wx.Frame):
 
    def __init__(self):
        wx.Frame.__init__(self, None, wx.ID_ANY, "Scorer",size=(750,70))
 
        # Add a panel so it looks the correct on all platforms
        panel = wx.Panel(self, wx.ID_ANY)
 
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.update, self.timer) 
        self.toggleBtn = wx.Button(panel, wx.ID_ANY, "Start")
        self.toggleBtn.Bind(wx.EVT_BUTTON, self.onToggle)        
        self.linkText = wx.TextCtrl(panel)
        
        hbox = wx.BoxSizer(wx.HORIZONTAL)        
        panel.SetSizer(hbox)
        
        hbox.Add(self.linkText, 4)
        hbox.Add((-1, 10))
 
        hbox.Add(self.toggleBtn)
        hbox.Add((-1, 10))
        
        self.link=''
        self.PreviousRuns = 0;
        self.PreviousWickets = 0;
        

    def onToggle(self, event):
        btnLabel = self.toggleBtn.GetLabel()
        self.link = self.linkText.GetValue()
        
        if btnLabel == "Start":
            print ("starting timer...")
            self.timer.Start(5000)
            self.toggleBtn.SetLabel("Stop")
        else:
            print ("timer stopped!")
            self.timer.Stop()
            self.toggleBtn.SetLabel("Start")
 
    def update(self, event):
        self.score = GetScore(self.link)        
        self.SetTitle(self.score)
        summary = self.score #Example score: 'PAK-W 70/8 (29.5 ov, batsman 0*, batsman2 2*, bowler 3/17)'
        summarySplit = summary.split(' ')
        score = summarySplit[1] 
        Runs = score.split('/')[0]
        Wickets = score.split('/')[1]
        
        if (int(Runs)-int(self.PreviousRuns)) >= 4:
            ShowNotification('Boundary !!!',self.score)
        
        if int(Wickets) - int(self.PreviousWickets) > 0:            
            ShowNotification('Wicket!!!',self.score)      
#            winsound.Beep(2500, 500)
        
        self.PreviousRuns = Runs;
        self.PreviousWickets = Wickets;        
        
def main():
   app = wx.App()
   frame = MyForm()
   frame.Show()
   app.MainLoop()


if __name__ == "__main__":
    main()
