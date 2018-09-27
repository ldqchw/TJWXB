#coding:utf-8
import os
import time
import pyttsx
import logging
import tldextract
from logging import handlers
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

class Logger(object):
    level_relations = {
        'debug':logging.DEBUG,
        'info':logging.INFO,
        'warning':logging.WARNING,
        'error':logging.ERROR,
        'crit':logging.CRITICAL
    }

    def __init__(self,filename,level='info',when='H',backCount=3,fmt='%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s'):
        self.logger = logging.getLogger(filename)
        format_str = logging.Formatter(fmt)
        self.logger.setLevel(self.level_relations.get(level))
        sh = logging.StreamHandler()
        sh.setFormatter(format_str)
        th = handlers.TimedRotatingFileHandler(filename=filename,when=when,backupCount=backCount,encoding='utf-8')
        th.setFormatter(format_str)
        self.logger.addHandler(sh)
        self.logger.addHandler(th)

class Dwszb:
    num = 1
    def __init__(self, fpath, upath,driver,timeout):
        self.fpath = fpath
        self.upath = upath
        self.driver = driver
        self.timeout = timeout
    def start(self):
        start = self.driver
        print "*"*50
        print "Dwszb "+start+" start...."
        if start == "firefox_start":
            self.firefox_start()
            self.read_url()
            self.driver_stop()
        elif start == "firefox_fb_start":
            self.firefox_fb_start()
            self.read_url()
            self.driver_stop()
        elif start == "chrome_start":
            self.chrome_start()
            self.read_url()
            self.driver_stop()
        elif start == "phantomJS_start":
            self.phantomJS_start()
            self.read_url()
            self.driver_stop()
        else:
            print "*" * 50
            print "Please input the following 4 strings for driver....\n1、firefox_start\n2、firefox_fb_start\n3、chrome_start\n4、phantomJS_start"
    def firefox_start(self):
        self.driver = webdriver.Firefox()
        self.driver.maximize_window()
    def chrome_start(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--user-agent=Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0); 360Spider")
        options.add_argument("--referer=https://www.so.com")
        self.driver = webdriver.Chrome(chrome_options=options)
        self.driver.maximize_window()
    def phantomJS_start(self):
        "selenium==2.48.0"
        dcap = dict(DesiredCapabilities.PHANTOMJS)
        dcap["phantomjs.page.settings.userAgent"] = (
            "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0); 360Spider")
        dcap["phantomjs.page.settings.referer"] = ("https://www.so.com")
        self.driver = webdriver.PhantomJS(desired_capabilities=dcap)
    def firefox_fb_start(self):
        profile_directory = r'C:\Users\Administrator\AppData\Roaming\Mozilla\Firefox\Profiles\xe5qiy45.default'
        profile = webdriver.FirefoxProfile(profile_directory)
        self.driver = webdriver.Firefox(profile)
        # time.sleep(3)
    def driver_stop(self):
        self.driver.quit()
        print "Dwszb Say..."
        self.say()
        print "Dwszb End...."
        print "*"*50
    def read_url(self):
        # i = 1
        f = open(self.upath, "r")
        for line in f:
            url = tldextract.extract(line)
            file = "{0}.{1}.{2}".format(url.subdomain, url.domain, url.suffix)
            self.screenshot(line,file)
        f.close()
    def screenshot(self,line,file):
        self.driver.set_page_load_timeout(self.timeout)
        self.driver.set_script_timeout(self.timeout)
        try:
            # self.driver.implicitly_wait(30)
            self.driver.get("http://" + line)
            # time.sleep(2)
        except:
            print "Access Timeout "+str(self.timeout)+"s:"+ line
            self.driver.execute_script('window.stop()')
        finally:
            self.driver.get_screenshot_as_file(
                self.if_file()+"\\" + str(Dwszb.num) + "-" + file + "-" + self.time_format() + ".png")
            print (self.if_file()+"\\" + str(Dwszb.num) + "-" + file + "-" + self.time_format() + ".png")
            Dwszb.num += 1
    def time_format(self):
        current_time = time.strftime('%Y%m%d-%H%M%S', time.localtime(time.time()))
        return current_time
    def if_file(self):
        dir_name = self.fpath+"ScreenShot-"+time.strftime('%m%d-%H', time.localtime(time.time()))
        if not os.path.isdir(dir_name):
            os.makedirs(dir_name)
        # print dir_name
        return dir_name
    def say(self):
        engine = pyttsx.init()
        engine.say(u'达沃斯重保单位的首页截图程序已执行完成！')
        engine.say(u'请工作人员检查保存的图片文件！')
        engine.say('The program has been completed.')
        engine.say('Please check the saved picture files.')
        engine.say('Please check the saved picture files.')
        engine.say('Please check the saved picture files.')
        engine.runAndWait()
if __name__ == '__main__':
    log = Logger('all.log',level='debug')
    Logger('error.log', level='error').logger.error('error')
    zb = Dwszb(".//",".//url.txt","firefox_start",15)
    zb.start()